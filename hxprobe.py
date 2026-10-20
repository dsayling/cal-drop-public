#!/usr/bin/env python3
"""cal-drop HxStore spike: scan HxStore.hxd (the new Outlook for Mac's
live store) and write hxstore-report.json.

COUNTS ONLY. The report holds record/format/object-type counts, how many
appointment-candidate records carry plausible timestamps, and the range
of those timestamps. It never extracts titles, subjects, names, or
bodies, so it is safe to share for the next fixing round — same rule as
probe.py.

Usage: python3 hxprobe.py [--path HxStore.hxd] [--out hxstore-report.json]
                          [--limit N]   (N>0 = quick partial scan)
"""

import argparse
import json
import sys
from datetime import datetime, timezone

import hxstore


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--path", default=None,
                    help="explicit HxStore.hxd path (default: discover)")
    ap.add_argument("--out", default="hxstore-report.json")
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    path = args.path
    if not path:
        paths = hxstore.find_hxstore_paths()
        if not paths:
            print("ERROR: no HxStore.hxd found under the Outlook profile "
                  "locations. Is this the new Outlook?", file=sys.stderr)
            return 2
        path = paths[0]

    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "tool": "cal-drop hxprobe v1 (spike)",
        "note": "Counts only. No titles, subjects, names, or bodies "
                "are extracted.",
        "store": hxstore.analyze_store(path, limit=args.limit),
    }
    with open(args.out, "w") as fh:
        json.dump(report, fh, indent=2)

    store = report["store"]
    print("Wrote %s" % args.out)
    print("- %s (%d bytes, magic ok: %s)" % (
        store["path"], store["size_bytes"], store["magic_ok"]))
    if store.get("error"):
        print("  ERROR: %s" % store["error"])
        return 3
    print("  pages: %d total, %d primary-store, %d secondary-store" % (
        store["pages"]["total"], store["pages"]["primary_store_pages"],
        store["pages"]["secondary_store_pages"]))
    rec = store["records"]
    print("  records scanned: %d (unique ids %d, failed %d)" % (
        rec["scanned"], rec["unique_record_ids"],
        rec["decompress_failed_or_invalid"]))
    print("  formats (payload u32): %s" % json.dumps(rec["formats"]))
    print("  raw stream formats: %s" % json.dumps(
        rec.get("raw_stream_formats")))
    print("  object types @44 (mixed records): %s" % json.dumps(
        rec["object_types_44_in_0x10013"]))
    print("  records containing class: %s" % json.dumps(
        rec["records_containing_class"]))
    appt = store["appointment_candidates"]
    print("  APPOINTMENT candidates: %d — with timestamps: %d, "
          "with 2+ plausible timestamps: %d" % (
              appt["count"], appt["with_sentinel_timestamps"],
              appt["with_2plus_plausible_timestamps"]))
    if appt["timestamp_range"]:
        print("  appointment timestamp range: %s .. %s" % (
            appt["timestamp_range"]["earliest"],
            appt["timestamp_range"]["latest"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
