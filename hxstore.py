"""HxStore.hxd reading for cal-drop — the spike.

The new Outlook for Mac (16.x) does not keep mail/calendar in the legacy
Outlook.sqlite at all. The live store is HxStore.hxd in the profile root:
Microsoft's undocumented "Cola" engine — 4096-byte pages, 8 slots of 512
bytes per page, records LZ4-block compressed, timestamps as .NET ticks.
Format layout follows the public reverse-engineering notes for
hxstore-decode (MIT; github.com/mitchell-johnson/hxstore-decode,
HXSTORE.md); this implementation is original, stdlib only.

This module is a SPIKE, not an exporter. It answers one question with
counts only: are there decodable Appointment records in the store, and
do they carry plausible timestamps (start/end candidates)? It extracts
no titles, subjects, names, or bodies, and its report contains none.

Reading is read-only; the upstream notes say the file is safe to read
while Outlook runs.
"""

import glob
import os
import re
import struct
from collections import Counter
from datetime import datetime, timedelta, timezone

MAGIC = b"Nostromoi"
PAGE_SIZE = 0x1000
SLOT_SIZE = 0x200
SLOTS_PER_PAGE = 8
SLOT_HEADER = 32
DATA_TYPE = 8
RECORD_ID_SIZE = 8

DOTNET_EPOCH = datetime(1, 1, 1, tzinfo=timezone.utc)
TICKS_PER_SECOND = 10_000_000
SENTINEL = b"\xff\x3f\x37\xf4\x75\x28\xca\x2b"

# Cola object types (from HxCore symbol analysis, HXSTORE.md section 14).
OBJECT_TYPES = {73: "Account", 77: "View", 104: "Calendar",
                107: "Appointment", 201: "MessageHeader",
                202: "DataReplication", 212: "AttachmentHeader",
                224: "Contact", 288: "SearchSession", 359: "Person"}

IPM_CLASSES = ["IPM.Note", "IPM.Appointment", "IPM.Schedule",
               "IPM.Meeting", "IPF.Note"]
IPM_UTF16 = {c: c.encode("utf-16-le") for c in IPM_CLASSES}
UTF16_RUN = re.compile(rb"(?:[\x20-\x7e]\x00){3,}")


def find_hxstore_paths():
    pattern = os.path.expanduser(
        "~/Library/Group Containers/UBF8T346G9.Office/Outlook/"
        "Outlook 15 Profiles/*/HxStore.hxd")
    found = sorted(set(glob.glob(pattern)))
    found.sort(key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0,
               reverse=True)
    return found


# ---------------------------------------------------------------- LZ4

def lz4_decompress_lenient(data, max_output):
    """Lenient LZ4 *block* decoder (raw blocks, no frame header).

    The standard library decoder rejects HxStore records whose final
    literal run is truncated; this one copies what is available and
    stops, per HXSTORE.md section 6. Pure Python; match copies are done
    in slices so the common cases are not byte-at-a-time.
    """
    out = bytearray()
    i = 0
    n = len(data)
    while i < n and len(out) < max_output:
        token = data[i]
        i += 1
        lit_len = token >> 4
        if lit_len == 15:
            while i < n:
                extra = data[i]
                i += 1
                lit_len += extra
                if extra != 255:
                    break
        take = min(lit_len, n - i, max_output - len(out))
        if take < 0:
            break
        out += data[i:i + take]
        i += take
        if take < lit_len or i >= n or len(out) >= max_output:
            break
        if i + 2 > n:
            break
        offset = struct.unpack_from("<H", data, i)[0]
        i += 2
        if offset == 0 or offset > len(out):
            break
        match_len = (token & 0x0F) + 4
        if (token & 0x0F) == 15:
            while i < n:
                extra = data[i]
                i += 1
                match_len += extra
                if extra != 255:
                    break
        match_len = min(match_len, max_output - len(out))
        start = len(out) - offset
        if offset >= match_len:
            out += out[start:start + match_len]
        else:
            # Overlapping match: the pattern repeats every `offset` bytes.
            pattern = bytes(out[start:])
            while match_len > 0:
                take = min(match_len, len(pattern))
                out += pattern[:take]
                match_len -= take
    return bytes(out)


# ------------------------------------------------------------ timestamps

def ticks_to_datetime(ticks):
    return DOTNET_EPOCH + timedelta(
        seconds=ticks // TICKS_PER_SECOND,
        microseconds=(ticks % TICKS_PER_SECOND) // 10)


def _ticks(dt):
    delta = dt - DOTNET_EPOCH
    return (delta.days * 86400 + delta.seconds) * TICKS_PER_SECOND


TICKS_MIN = _ticks(datetime(2015, 1, 1, tzinfo=timezone.utc))
TICKS_MAX = _ticks(datetime(2035, 1, 1, tzinfo=timezone.utc))


def sentinel_timestamps(blob):
    """Plausible .NET-tick values from sentinel-bracketed blocks.

    Layout (HXSTORE.md section 12): an int64 timestamp, then the
    sentinel, then another int64 (the display time), then more
    sentinels. Returns the plausible tick values found around every
    sentinel occurrence: the int64 just before it and just after it.
    """
    return [value for _, value in sentinel_timestamps_with_offsets(blob)]


def sentinel_timestamps_with_offsets(blob):
    """Same scan as sentinel_timestamps, keeping each value's offset."""
    found = []
    start = 0
    while True:
        pos = blob.find(SENTINEL, start)
        if pos < 0:
            return found
        for cand_pos in (pos - 8, pos + 8):
            if 0 <= cand_pos and cand_pos + 8 <= len(blob):
                value = struct.unpack_from("<q", blob, cand_pos)[0]
                if TICKS_MIN <= value <= TICKS_MAX:
                    found.append((cand_pos, value))
        start = pos + 1


ASCII_RUN = re.compile(rb"[\x20-\x7e]{5,}")


def utf16_strings_with_offsets(blob):
    """[(offset, text)] for UTF-16LE printable runs of 3+ characters."""
    out = []
    for m in UTF16_RUN.finditer(blob):
        try:
            text = m.group().decode("utf-16-le")
        except UnicodeDecodeError:
            continue
        out.append((m.start(), text))
    return out


def ascii_strings_with_offsets(blob):
    """[(offset, text)] for ASCII printable runs of 5+ characters."""
    return [(m.start(), m.group().decode("ascii"))
            for m in ASCII_RUN.finditer(blob)]


def all_plausible_ticks(blob):
    """[(offset, ticks)] for every int64 in the plausible .NET-tick
    range, scanning every byte offset. Diagnostic use on single records
    (field mapping), not a whole-store scan."""
    out = []
    for off in range(0, len(blob) - 7):
        value = struct.unpack_from("<q", blob, off)[0]
        if TICKS_MIN <= value <= TICKS_MAX:
            out.append((off, value))
    return out


# ------------------------------------------------------------ records

def discover_store_ids(data):
    """Frequency-scan page headers: the primary store id is the most
    common 8-byte value at page+8 among pages whose type (page+16) is 8.
    Returns (primary, secondary, Counter)."""
    counts = Counter()
    pages = len(data) // PAGE_SIZE
    for page in range(pages):
        off = page * PAGE_SIZE
        if struct.unpack_from("<I", data, off + 16)[0] != DATA_TYPE:
            continue
        sid = data[off + 8:off + 16]
        if sid != b"\x00" * 8:
            counts[sid] += 1
    ranked = counts.most_common()
    primary = ranked[0][0] if ranked else None
    secondary = ranked[1][0] if len(ranked) > 1 else None
    return primary, secondary, counts


def iter_records(data, primary_sid):
    """Yield (record_id, decompressed_payload, size_a, size_b,
    file_offset, raw_stream_format, slot_unknown) for every data record
    in the file, following multi-slot records by file offset so their
    continuation slots are not re-read as headers. raw_stream_format is
    the uint32 at the start of the compressed stream (the value upstream
    hxdecode reports as the record format type); slot_unknown is the
    slot header's trailing uint32 (observed values 2, 4, 6)."""
    total = len(data)
    skip_until = 0
    failed = 0
    pages = total // PAGE_SIZE
    for page in range(pages):
        for slot in range(SLOTS_PER_PAGE):
            off = page * PAGE_SIZE + slot * SLOT_SIZE
            if off < skip_until:
                continue
            if data[off + 8:off + 16] != primary_sid:
                continue
            typ, size_a, size_b, unknown = struct.unpack_from(
                "<IIII", data, off + 16)
            if typ != DATA_TYPE:
                continue
            if size_a < RECORD_ID_SIZE + 1 or size_b < RECORD_ID_SIZE:
                continue
            if size_b > 512 * 1024 * 1024 or off + SLOT_HEADER + size_a > total:
                failed += 1
                continue
            raw = data[off + SLOT_HEADER: off + SLOT_HEADER + size_a]
            record_id = struct.unpack_from("<Q", raw, 0)[0]
            raw_fmt = struct.unpack_from("<I", raw, RECORD_ID_SIZE)[0] \
                if len(raw) >= RECORD_ID_SIZE + 4 else None
            payload = lz4_decompress_lenient(
                raw[RECORD_ID_SIZE:], size_b - RECORD_ID_SIZE)
            if not payload:
                failed += 1
                continue
            span = SLOT_HEADER + size_a
            skip_until = off + ((span + SLOT_SIZE - 1) // SLOT_SIZE) * SLOT_SIZE
            yield record_id, payload, size_a, size_b, off, raw_fmt, unknown
    return failed


def analyze_record(payload):
    fmt = struct.unpack_from("<I", payload, 0)[0] if len(payload) >= 4 \
        else None
    obj44 = None
    # The mixed/metadata record family (upstream format 0x10013) is
    # recognised in the *decompressed* payload by its 64-byte header:
    # zeros at 0-7 and the Cola store tag 0x0570 at offset 18. Its
    # ObjectType sits at offset 44 (HXSTORE.md section 8).
    if len(payload) >= 46 and payload[0:8] == b"\x00" * 8 and \
            struct.unpack_from("<H", payload, 18)[0] == 0x0570:
        obj44 = struct.unpack_from("<H", payload, 44)[0]
    classes = [c for c in IPM_CLASSES if IPM_UTF16[c] in payload]
    ticks = sentinel_timestamps(payload)
    return {
        "format": fmt,
        "object_type_44": obj44,
        "classes": classes,
        "ticks": ticks,
        "utf16_strings": len(UTF16_RUN.findall(payload)),
        "size": len(payload),
    }


def is_appointment_candidate(info):
    if info["object_type_44"] in (104, 107):
        return True
    return any(c in info["classes"]
               for c in ("IPM.Appointment", "IPM.Schedule", "IPM.Meeting"))


def _size_stats(values):
    if not values:
        return None
    ordered = sorted(values)
    return {"min": ordered[0], "median": ordered[len(ordered) // 2],
            "max": ordered[-1]}


def analyze_store(path, limit=0):
    """Full spike scan of one HxStore.hxd. Counts only — see module
    docstring. `limit` > 0 stops after that many records (quick pass)."""
    with open(path, "rb") as fh:
        data = fh.read()
    report = {
        "path": path,
        "size_bytes": len(data),
        "magic_ok": data[:9] == MAGIC + b"\x00" or data[:9] == MAGIC,
    }
    if len(data) >= 0x30:
        report["header"] = {
            "logical_size": struct.unpack_from("<Q", data, 0x10)[0],
            "record_count_field": struct.unpack_from("<I", data, 0x28)[0],
        }
    pages = len(data) // PAGE_SIZE
    primary, secondary, sid_counts = discover_store_ids(data)
    report["pages"] = {
        "total": pages,
        "primary_store_pages": sid_counts.get(primary, 0) if primary else 0,
        "secondary_store_pages": sid_counts.get(secondary, 0)
        if secondary else 0,
        "distinct_store_ids": len(sid_counts),
    }
    report["store_ids"] = {
        "primary": primary.hex() if primary else None,
        "secondary": secondary.hex() if secondary else None,
    }
    if primary is None:
        report["error"] = "no data-page store id found"
        return report

    formats = Counter()
    raw_formats = Counter()
    objtypes = Counter()
    class_counts = Counter()
    seen_ids = set()
    records = 0
    gen = iter_records(data, primary)
    failed = 0
    appt = []
    while True:
        try:
            record_id, payload, size_a, size_b, _off, raw_fmt, _u = \
                next(gen)
        except StopIteration as stop:
            failed = stop.value or 0
            break
        records += 1
        seen_ids.add(record_id)
        info = analyze_record(payload)
        formats[info["format"]] += 1
        if raw_fmt is not None:
            raw_formats[raw_fmt] += 1
        if info["object_type_44"] is not None:
            objtypes[info["object_type_44"]] += 1
        for c in info["classes"]:
            class_counts[c] += 1
        if is_appointment_candidate(info):
            appt.append(info)
        if limit and records >= limit:
            break

    report["records"] = {
        "scanned": records,
        "unique_record_ids": len(seen_ids),
        "decompress_failed_or_invalid": failed,
        "formats": {"0x%X" % k: v for k, v in formats.most_common()
                    if k is not None},
        "raw_stream_formats": {"0x%X" % k: v for k, v in
                               raw_formats.most_common() if k is not None},
        "object_types_44_in_0x10013": {
            ("%d (%s)" % (k, OBJECT_TYPES.get(k, "?"))): v
            for k, v in objtypes.most_common()},
        "records_containing_class": dict(class_counts.most_common()),
    }

    all_ticks = [t for info in appt for t in info["ticks"]]
    report["appointment_candidates"] = {
        "count": len(appt),
        "by_object_type_44": dict(Counter(
            str(i["object_type_44"]) for i in appt).most_common()),
        "by_format": {"0x%X" % k: v for k, v in Counter(
            i["format"] for i in appt).most_common() if k is not None},
        "with_sentinel_timestamps": sum(1 for i in appt if i["ticks"]),
        "with_2plus_plausible_timestamps": sum(
            1 for i in appt if len(set(i["ticks"])) >= 2),
        "plausible_timestamp_values": len(all_ticks),
        "timestamp_range": None,
        "decompressed_size": _size_stats([i["size"] for i in appt]),
        "utf16_strings_per_record": _size_stats(
            [i["utf16_strings"] for i in appt]),
    }
    if all_ticks:
        report["appointment_candidates"]["timestamp_range"] = {
            "earliest": ticks_to_datetime(min(all_ticks)).isoformat(),
            "latest": ticks_to_datetime(max(all_ticks)).isoformat(),
        }
    return report
