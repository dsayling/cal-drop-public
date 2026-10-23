"""Tests for the HxStore exporter: a synthetic store carrying an
appointment-shaped Cola record (signature header, adjacent start/end
ticks, block-ordered strings) must export exactly that event.

    python3 -m unittest discover -s tests -v
"""

import os
import struct
import sys
import tempfile
import unittest
from datetime import datetime, timedelta, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import hxexport  # noqa: E402
import hxstore  # noqa: E402
import outlook_db as odb  # noqa: E402
from test_hxstore import build_store  # noqa: E402


def put_utf16(buf, off, text):
    raw = text.encode("utf-16-le")
    buf[off:off + len(raw)] = raw


def put_ticks(buf, off, value):
    struct.pack_into("<q", buf, off, value)


def make_appointment_payload(start_dt, end_dt):
    """Lay out a payload the way the real test-event record is laid
    out: header (u32 1 @0, size-4 @4, type 107 @29), a lone reminder
    tick, the start/end pair, timezone strings, hex blob, title,
    organizer, location — then the pair and title repeated (block 2)."""
    buf = bytearray(2200)
    struct.pack_into("<I", buf, 0, 1)
    struct.pack_into("<I", buf, 4, len(buf) - 4)
    struct.pack_into("<H", buf, 29, 107)
    start_t = hxstore._ticks(start_dt.astimezone(timezone.utc))
    end_t = hxstore._ticks(end_dt.astimezone(timezone.utc))
    put_ticks(buf, 483, start_t - 15 * 60 * hxstore.TICKS_PER_SECOND)
    put_ticks(buf, 603, start_t)
    put_ticks(buf, 611, end_t)
    put_utf16(buf, 700, "US/Mountain")
    put_utf16(buf, 900, "AB12" * 28)
    put_utf16(buf, 1126, "EST5EDT")
    put_utf16(buf, 1150, "CALDROP-TEST-7F3A")
    put_utf16(buf, 1250, "Example, Alex (A.)")
    put_utf16(buf, 1350, "alex@example.com")
    put_utf16(buf, 1500, "Test Room")
    put_ticks(buf, 1700, start_t)
    put_ticks(buf, 1708, end_t)
    put_utf16(buf, 1800, "CALDROP-TEST-7F3A")
    return bytes(buf)


def make_mail_payload(start_dt, end_dt):
    """Decoy: plausible adjacent ticks, but no appointment signature."""
    buf = bytearray(800)
    struct.pack_into("<I", buf, 0, 3)
    put_ticks(buf, 100, hxstore._ticks(start_dt.astimezone(timezone.utc)))
    put_ticks(buf, 108, hxstore._ticks(end_dt.astimezone(timezone.utc)))
    put_utf16(buf, 300, "IPM.Note")
    put_utf16(buf, 400, "Some mail subject")
    return bytes(buf)


class HxExportTests(unittest.TestCase):
    def setUp(self):
        local = odb.LOCAL_TZ
        day = datetime.now(local) + timedelta(days=1)
        self.start = day.replace(hour=15, minute=0, second=0,
                                 microsecond=0)
        self.end = day.replace(hour=16, minute=30, second=0,
                               microsecond=0)
        self.appt = make_appointment_payload(self.start, self.end)

    def test_units(self):
        self.assertTrue(hxexport.is_appointment(self.appt))
        pair = hxexport.find_start_end(self.appt)
        self.assertEqual(pair, (
            hxstore._ticks(self.start.astimezone(timezone.utc)),
            hxstore._ticks(self.end.astimezone(timezone.utc))))
        self.assertEqual(hxexport.extract_title(self.appt),
                         "CALDROP-TEST-7F3A")

    def test_merged_tz_title_is_peeled(self):
        buf = bytearray(self.appt)
        put_utf16(buf, 1150, "America/BahiaTeam 3.6 - Project Review")
        self.assertEqual(hxexport.extract_title(bytes(buf)),
                         "Team 3.6 - Project Review")
        self.assertTrue(hxexport._is_tz("EST5EDT"))
        self.assertTrue(hxexport._is_tz("MST7MDT"))
        self.assertTrue(hxexport._is_tz("tzone://Microsoft/Utc"))
        self.assertFalse(hxexport._is_tz("Example/Widgets - Weekly Sync"))

    def test_exception_instance_body_line_loses_to_repeated_subject(self):
        """Regression: an edited (exception) instance of a recurring
        meeting stores the first body/agenda line as the FIRST content
        string after the hex timezone blob, with the real subject
        appearing later — repeated twice (before and after the
        organizer). The title must be the repeated subject, not the
        body line. Shape mirrors live record 2945126656219980196
        (2026-10-23 09:05 America/Denver)."""
        buf = bytearray(2200)
        struct.pack_into("<I", buf, 0, 1)
        struct.pack_into("<I", buf, 4, len(buf) - 4)
        struct.pack_into("<H", buf, 29, 107)
        start_t = hxstore._ticks(self.start.astimezone(timezone.utc))
        end_t = hxstore._ticks(self.end.astimezone(timezone.utc))
        put_ticks(buf, 483, start_t - 15 * 60 * hxstore.TICKS_PER_SECOND)
        put_ticks(buf, 603, start_t)
        put_ticks(buf, 611, end_t)
        put_utf16(buf, 700, "AB12" * 28)          # hex blob, first string
        put_utf16(buf, 1100, "Eastern Standard Time")
        # Body/agenda first line — WRONG title candidate, first content string.
        put_utf16(buf, 1300, "Updated 10/23 Agenda. Unfortunately, the demo needs to be delayed")
        put_utf16(buf, 1500, "Purpose:")
        put_utf16(buf, 1600, "Dupuis, Michael (M.C.)")
        put_utf16(buf, 1700, "mdupuis9@ford.com")
        # Real subject — appears twice, as in every observed instance.
        put_utf16(buf, 1800, "NEXT Overall Forum - Friday")
        put_utf16(buf, 1900, "Ayling, Drew (D.)")
        put_utf16(buf, 2000, "NEXT Overall Forum - Friday")
        put_ticks(buf, 2100, start_t)
        put_ticks(buf, 2108, end_t)
        self.assertEqual(hxexport.extract_title(bytes(buf)),
                         "NEXT Overall Forum - Friday")

    def test_store_end_to_end(self):
        tmp = tempfile.mkdtemp(prefix="caldrop-hxexp-")
        path = os.path.join(tmp, "HxStore.hxd")
        build_store(path, [self.appt,
                           make_mail_payload(self.start, self.end)])
        snapshot, diagnostics, code = hxexport.build_snapshot_hx(path, 14)
        self.assertIsNotNone(snapshot, diagnostics)
        self.assertEqual(code, 0)
        self.assertEqual(len(snapshot["events"]), 1, snapshot["events"])
        ev = snapshot["events"][0]
        self.assertEqual(ev["title"], "CALDROP-TEST-7F3A")
        self.assertTrue(ev["start"].startswith(
            self.start.strftime("%Y-%m-%dT15:00")), ev)
        self.assertTrue(ev["end"].startswith(
            self.end.strftime("%Y-%m-%dT16:30")), ev)
        self.assertFalse(ev["all_day"])
        self.assertEqual(ev["show_as"], "busy")
        self.assertEqual(
            diagnostics["stats"]["signature_matches"], 1)


if __name__ == "__main__":
    unittest.main()
