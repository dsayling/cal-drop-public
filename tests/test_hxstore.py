"""Tests for the HxStore spike: build synthetic HxStore.hxd files using
the documented layout (HXSTORE.md) and run the real scanner over them.

    python3 -m unittest discover -s tests -v
"""

import json
import os
import struct
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import hxstore  # noqa: E402

PRIMARY = bytes.fromhex("1122334455667788")
SECONDARY = bytes.fromhex("99aabbccddeeff00")


def lz4_literal_block(payload):
    """Encode payload as one literal-only LZ4 block sequence (valid LZ4:
    a single sequence whose literal length uses extension bytes, ending
    at end-of-stream)."""
    n = len(payload)
    if n < 15:
        return bytes([n << 4]) + payload
    extra = n - 15
    ext = bytearray()
    while extra >= 255:
        ext.append(255)
        extra -= 255
    ext.append(extra)
    return bytes([0xF0]) + bytes(ext) + payload


def ticks(dt):
    return hxstore._ticks(dt)


def sentinel_block(sync_ticks, display_ticks):
    return (struct.pack("<q", sync_ticks) + hxstore.SENTINEL +
            struct.pack("<q", display_ticks) + hxstore.SENTINEL * 3)


def make_payload(fmt, obj44=None, classes=(), tick_pairs=(), filler=b""):
    buf = bytearray(struct.pack("<I", fmt))
    if fmt == 0x10013:
        buf += b"\x00" * (44 - len(buf))
        buf += struct.pack("<H", obj44 or 0)
        buf += b"\x00" * 22
    for cls in classes:
        buf += cls.encode("utf-16-le") + b"\x00\x00"
    buf += filler
    for sync_t, disp_t in tick_pairs:
        buf += sentinel_block(sync_t, disp_t)
    return bytes(buf)


def build_store(path, payloads):
    """Assemble a minimal HxStore: header page, one secondary-store page,
    then data records laid out at slot boundaries (multi-slot records
    span slots/pages contiguously, like the real file)."""
    out = bytearray()
    header = bytearray(hxstore.PAGE_SIZE)
    header[0:9] = b"Nostromoi"
    struct.pack_into("<I", header, 0x28, len(payloads))
    out += header

    secondary = bytearray(hxstore.PAGE_SIZE)
    struct.pack_into("<Q", secondary, 0, 123)          # hash
    secondary[8:16] = SECONDARY
    struct.pack_into("<I", secondary, 16, 8)           # type
    struct.pack_into("<I", secondary, 20, 0x200)
    out += secondary

    for idx, payload in enumerate(payloads, start=1000):
        comp = lz4_literal_block(payload)
        size_a = 8 + len(comp)
        size_b = 8 + len(payload)
        slot = bytearray()
        slot += struct.pack("<Q", 42 + idx)            # hash
        slot += PRIMARY
        slot += struct.pack("<III", 8, size_a, size_b)
        slot += struct.pack("<I", 2)                   # unknown field
        assert len(slot) == hxstore.SLOT_HEADER
        slot += struct.pack("<Q", idx)                 # record id
        slot += comp
        out += slot
        # One record per page region: keeps the primary store id the
        # most common page-header id, as in real files.
        pad = (-len(out)) % hxstore.PAGE_SIZE
        out += b"\x00" * pad
    with open(path, "wb") as fh:
        fh.write(bytes(out))


class Lz4Tests(unittest.TestCase):
    def test_literal_roundtrip(self):
        payload = b"hello hxstore " * 40
        decoded = hxstore.lz4_decompress_lenient(
            lz4_literal_block(payload), len(payload))
        self.assertEqual(decoded, payload)

    def test_match_copy(self):
        # seq: 4 literals "abcd", match offset 4 len 4, then empty token
        stream = b"\x40abcd\x04\x00\x00"
        decoded = hxstore.lz4_decompress_lenient(stream, 8)
        self.assertEqual(decoded, b"abcdabcd")

    def test_overlapping_match(self):
        # literals "ab", then match offset 2 len 6 repeats the pattern
        stream = b"\x22ab\x02\x00\x00"
        decoded = hxstore.lz4_decompress_lenient(stream, 8)
        self.assertEqual(decoded, b"abababab")


class SpikeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="caldrop-hx-")
        self.path = os.path.join(self.tmp, "HxStore.hxd")
        sync_t = ticks(datetime(2026, 10, 1, tzinfo=timezone.utc))
        disp_t = ticks(datetime(2026, 10, 21, 15, 0, tzinfo=timezone.utc))
        appt = make_payload(
            0x10013, obj44=107, classes=("IPM.Appointment",),
            tick_pairs=((sync_t, disp_t),),
            filler="SECRET MEETING TITLE".encode("utf-16-le"))
        mail = make_payload(
            0x03B0, classes=("IPM.Note",),
            filler="MAIL ONLY SECRET".encode("utf-16-le"))
        big_appt = make_payload(
            0x10013, obj44=104, classes=("IPM.Schedule",),
            tick_pairs=((sync_t, disp_t),),
            filler=("x".encode("utf-16-le")) * 500)
        build_store(self.path, [appt, mail, big_appt])
        self.report = hxstore.analyze_store(self.path)

    def test_counts(self):
        rec = self.report["records"]
        self.assertTrue(self.report["magic_ok"])
        self.assertEqual(rec["scanned"], 3)
        self.assertEqual(rec["unique_record_ids"], 3)
        self.assertEqual(rec["formats"]["0x10013"], 2)
        self.assertEqual(rec["formats"]["0x3B0"], 1)
        self.assertEqual(rec["records_containing_class"]["IPM.Note"], 1)
        self.assertEqual(
            self.report["pages"]["primary_store_pages"], 3)

    def test_appointment_candidates(self):
        appt = self.report["appointment_candidates"]
        self.assertEqual(appt["count"], 2)
        self.assertEqual(appt["with_sentinel_timestamps"], 2)
        self.assertEqual(appt["with_2plus_plausible_timestamps"], 2)
        self.assertIn("2026-10-01", appt["timestamp_range"]["earliest"])
        self.assertIn("2026-10-21", appt["timestamp_range"]["latest"])

    def test_hxfind_dumps_only_the_matched_record(self):
        out = os.path.join(self.tmp, "hxfind-report.json")
        proc = subprocess.run(
            [sys.executable, os.path.join(ROOT, "hxfind.py"),
             "SECRET MEETING TITLE", "--path", self.path, "--out", out],
            capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        with open(out) as fh:
            text = fh.read()
        self.assertIn("SECRET MEETING TITLE", text)
        self.assertNotIn("MAIL ONLY SECRET", text)
        report = json.loads(text)
        self.assertEqual(report["matching_record_copies"], 1)
        match = report["matches"][0]
        self.assertEqual(len(match["hex"]),
                         match["decompressed_size"] * 2)
        self.assertTrue(match["title_match_offsets_utf16"])
        isos = [t["iso"] for t in match["all_plausible_ticks"]]
        self.assertTrue(any(i.startswith("2026-10-21T15:00")
                            for i in isos), isos)

    def test_cli_report_contains_no_titles(self):
        out = os.path.join(self.tmp, "hxstore-report.json")
        proc = subprocess.run(
            [sys.executable, os.path.join(ROOT, "hxprobe.py"),
             "--path", self.path, "--out", out],
            capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        with open(out) as fh:
            text = fh.read()
        self.assertNotIn("SECRET MEETING TITLE", text)
        report = json.loads(text)
        self.assertEqual(
            report["store"]["appointment_candidates"]["count"], 2)


if __name__ == "__main__":
    unittest.main()
