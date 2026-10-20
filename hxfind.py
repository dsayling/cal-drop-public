#!/usr/bin/env python3
"""cal-drop field-mapping tool: find HxStore records containing a search
string and dump their full internal structure to hxfind-report.json.

This is a DIFFERENTIAL-ANALYSIS tool, meant to be run against a
synthetic test event you created on purpose (unique title, known start
and end). It is the one cal-drop tool whose report DOES contain record
contents — but only for records that match the search string, so the
only event data in the report is the test event's. Do not run it with a
real meeting title unless you mean to share that meeting's record.

For each matching record the report lists, with byte offsets:
  - every UTF-16LE and ASCII string in the record
  - every plausible .NET-tick timestamp (sentinel-adjacent ones first,
    then a full int64 scan), decoded to ISO UTC
Knowing the test event's real start/end in advance is what lets the
field mapping be read off the dump: the offsets whose decoded values
equal the known times are the start/end fields, and the title's string
index/offset shows where titles live.

Usage: python3 hxfind.py "CALDROP-TEST-7F3A" [--path HxStore.hxd]
                         [--out hxfind-report.json] [--max 5]
"""

import argparse
import json
import sys
from datetime import datetime, timezone

import hxstore

MAX_STRINGS = 400
MAX_TICKS = 600


HEX_CAP = 16384


def describe_record(record_id, payload, size_a, size_b, offset, raw_fmt,
                    needle_utf16, slot_unknown=None):
    strings16 = hxstore.utf16_strings_with_offsets(payload)
    strings_a = hxstore.ascii_strings_with_offsets(payload)
    sentinel = hxstore.sentinel_timestamps_with_offsets(payload)
    all_ticks = hxstore.all_plausible_ticks(payload)
    return {
        "record_id": record_id,
        "file_offset": offset,
        "slot_unknown_field": slot_unknown,
        "size_a": size_a,
        "size_b": size_b,
        "decompressed_size": len(payload),
        "hex": payload.hex() if len(payload) <= HEX_CAP else None,
        "hex_omitted_reason": None if len(payload) <= HEX_CAP else
        "record larger than %d bytes" % HEX_CAP,
        "raw_stream_format": "0x%X" % raw_fmt if raw_fmt is not None
        else None,
        "payload_first_u32": struct_fmt(payload),
        "title_match_offsets_utf16": [
            off for off, text in strings16 if needle_utf16 in text],
        "utf16_strings": [
            {"index": i, "offset": off, "text": text}
            for i, (off, text) in enumerate(strings16[:MAX_STRINGS])],
        "utf16_string_count": len(strings16),
        "utf16_strings_truncated": len(strings16) > MAX_STRINGS,
        "ascii_strings": [
            {"index": i, "offset": off, "text": text}
            for i, (off, text) in enumerate(strings_a[:MAX_STRINGS])],
        "ascii_string_count": len(strings_a),
        "sentinel_timestamps": [
            {"offset": off, "ticks": value,
             "iso": hxstore.ticks_to_datetime(value).isoformat()}
            for off, value in sentinel[:MAX_TICKS]],
        "all_plausible_ticks": [
            {"offset": off, "ticks": value,
             "iso": hxstore.ticks_to_datetime(value).isoformat()}
            for off, value in all_ticks[:MAX_TICKS]],
        "all_plausible_ticks_truncated": len(all_ticks) > MAX_TICKS,
        "section_marker_offsets": marker_offsets(payload),
    }


def struct_fmt(payload):
    import struct
    return ("0x%X" % struct.unpack_from("<I", payload, 0)[0]
            if len(payload) >= 4 else None)


def marker_offsets(payload):
    marker = b"\x00\x01\x00\x00\x00\x00\x01"
    out = []
    start = 0
    while True:
        pos = payload.find(marker, start)
        if pos < 0:
            return out[:100]
        out.append(pos)
        start = pos + 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("search", help="unique string from a synthetic test "
                                   "event you created for this purpose")
    ap.add_argument("--path", default=None)
    ap.add_argument("--out", default="hxfind-report.json")
    ap.add_argument("--max", type=int, default=5,
                    help="max matching records to dump in full")
    args = ap.parse_args()

    path = args.path
    if not path:
        paths = hxstore.find_hxstore_paths()
        if not paths:
            print("ERROR: no HxStore.hxd found.", file=sys.stderr)
            return 2
        path = paths[0]

    with open(path, "rb") as fh:
        data = fh.read()
    primary, _secondary, _counts = hxstore.discover_store_ids(data)
    if primary is None:
        print("ERROR: no data-page store id found.", file=sys.stderr)
        return 3

    needle16 = args.search.encode("utf-16-le")
    needle_a = args.search.encode("ascii", "ignore")
    matches = []
    copies = 0
    scanned = 0
    gen = hxstore.iter_records(data, primary)
    while True:
        try:
            record_id, payload, size_a, size_b, offset, raw_fmt, \
                unknown = next(gen)
        except StopIteration:
            break
        scanned += 1
        if needle16 not in payload and (
                not needle_a or needle_a not in payload):
            continue
        copies += 1
        if len(matches) < args.max:
            matches.append(describe_record(
                record_id, payload, size_a, size_b, offset, raw_fmt,
                args.search, slot_unknown=unknown))

    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "tool": "cal-drop hxfind v2",
        "warning": "This report contains the full string/timestamp "
                   "structure of records matching the search string "
                   "ONLY. It is meant for a synthetic test event; "
                   "anything else that matched is in here too.",
        "search": args.search,
        "store_path": path,
        "records_scanned": scanned,
        "matching_record_copies": copies,
        "matches_dumped": len(matches),
        "matches": matches,
    }
    with open(args.out, "w") as fh:
        json.dump(report, fh, indent=2)
    print("Wrote %s" % args.out)
    print("Scanned %d records; %d copies match %r; dumped %d." % (
        scanned, copies, args.search, len(matches)))
    for m in matches:
        print("- record %d @%d: %d bytes, %d utf16 strings, "
              "%d sentinel timestamps, %d plausible ticks" % (
                  m["record_id"], m["file_offset"],
                  m["decompressed_size"], m["utf16_string_count"],
                  len(m["sentinel_timestamps"]),
                  len(m["all_plausible_ticks"])))
    return 0 if copies else 4


if __name__ == "__main__":
    sys.exit(main())
