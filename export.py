#!/usr/bin/env python3
"""cal-drop exporter: read the local Outlook for Mac SQLite cache and write
calendar.json — title, start, end, all-day, show-as only. No attendees, no
bodies, no locations, no meeting links ever leave the machine; the filter
is applied here, before anything is written.

Built on assumptions (see ASSUMPTIONS.md): adaptive table/column discovery and
sample-based date-decoder selection, with every choice recorded under
"source" in the output so a wrong guess is visible, not silent.

Usage: python3 export.py [--db PATH] [--days 14] [--out calendar.json]
Exit codes: 0 ok, 2 no database found, 3 no usable event table,
4 event table found but zero events decoded (see diagnostics in output).
"""

import argparse
import json
import os
import shutil
import sys
from datetime import datetime, time, timedelta, timezone

import outlook_db as odb

PRIVATE_TITLE = "Private appointment"


def build_snapshot(db_path, days):
    diagnostics = {"db_path": db_path}
    tmpdir, conn = odb.open_db_copy(db_path)
    try:
        tables = odb.list_tables(conn)
        cands = odb.table_candidates(tables)
        top = cands[0] if cands else None
        if not top or not top["columns"]["subject"] or \
                not top["columns"]["start"] or top["score"] < 5:
            diagnostics["candidates"] = cands[:5]
            return None, diagnostics, 3
        cols = top["columns"]
        diagnostics.update({
            "table": top["table"], "columns": cols,
            "table_row_count": top["row_count"],
        })
        rows = conn.execute('SELECT * FROM "%s"' % top["table"]).fetchall()
        diagnostics["rows_read"] = len(rows)

        raw_starts = [r[cols["start"]] for r in rows]
        decoder = odb.choose_decoder(raw_starts)
        diagnostics["date_decoder"] = decoder

        now_local = datetime.now(odb.LOCAL_TZ)
        window_start_local = datetime.combine(
            now_local.date(), time.min, tzinfo=odb.LOCAL_TZ)
        window_start = window_start_local.astimezone(timezone.utc)
        window_end = window_start + timedelta(days=days)

        events = []
        stats = {"skipped_no_start": 0, "skipped_cancelled": 0,
                 "recurring_masters": 0, "recurrence_unparsed": 0,
                 "private_redacted": 0}
        seen = set()
        for row in rows:
            if cols["cancelled"] and odb.is_truthy(row[cols["cancelled"]]):
                stats["skipped_cancelled"] += 1
                continue
            start = odb.decode_date(row[cols["start"]], decoder)
            if start is None:
                stats["skipped_no_start"] += 1
                continue
            end = None
            if cols["end"]:
                end = odb.decode_date(row[cols["end"]], decoder)
            all_day = bool(cols["all_day"] and
                           odb.is_truthy(row[cols["all_day"]]))
            if end is None or end <= start:
                end = start + (timedelta(days=1) if all_day
                               else timedelta(hours=1))
            title = row[cols["subject"]]
            title = "" if title is None else str(title).strip()
            if cols["sensitivity"] and odb.is_private(
                    row[cols["sensitivity"]], cols["sensitivity"]):
                title = PRIVATE_TITLE
                stats["private_redacted"] += 1
            if not title:
                title = "(no title)"
            show_as = odb.map_show_as(row[cols["show_as"]]) \
                if cols["show_as"] else None
            uid = str(row[cols["uid"]]) if cols["uid"] and \
                row[cols["uid"]] is not None else None

            starts = [start]
            if cols["recurrence"] and row[cols["recurrence"]] not in (
                    None, "", 0, b""):
                rule = odb.parse_rrule(row[cols["recurrence"]])
                stats["recurring_masters"] += 1
                if rule:
                    starts = odb.expand_rrule(
                        start, rule, window_start, window_end) or []
                else:
                    stats["recurrence_unparsed"] += 1
                    starts = [start] if window_start <= start < window_end \
                        else []
            else:
                starts = [start] if window_start <= start < window_end \
                    else []
            duration = end - start
            for occ in starts:
                key = (uid, occ.isoformat()) if uid else \
                    (title, occ.isoformat())
                if key in seen:
                    continue
                seen.add(key)
                events.append({
                    "title": title,
                    "start": occ.astimezone(odb.LOCAL_TZ).isoformat(),
                    "end": (occ + duration).astimezone(
                        odb.LOCAL_TZ).isoformat(),
                    "all_day": all_day,
                    "show_as": show_as,
                })
        events.sort(key=lambda e: (e["start"], e["title"]))
        diagnostics["stats"] = stats
        diagnostics["events_out"] = len(events)
        snapshot = {
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "timezone": str(odb.LOCAL_TZ),
            "window_days": days,
            "source": {
                "app": "outlook-for-mac-local-sqlite (assumed)",
                "profile_dir": os.path.basename(os.path.dirname(
                    os.path.dirname(db_path))),
                "table": top["table"],
                "columns": cols,
                "date_decoder": decoder,
                "stats": stats,
            },
            "events": events,
        }
        code = 0 if events else 4
        return snapshot, diagnostics, code
    finally:
        conn.close()
        shutil.rmtree(tmpdir, ignore_errors=True)


def build_snapshot_auto(config, days, db_override=None):
    """Prefer the live HxStore (the new Outlook's store); fall back to
    the legacy SQLite cache. Returns (snapshot|None, diagnostics,
    code), same contract as build_snapshot."""
    if db_override:
        return build_snapshot(db_override, days)
    import hxstore
    hx_paths = hxstore.find_hxstore_paths()
    if hx_paths:
        import hxexport
        return hxexport.build_snapshot_hx(hx_paths[0], days)
    db_path = config.get("db_path")
    if not db_path:
        paths = odb.find_db_paths()
        if not paths:
            return None, {"error": "no HxStore.hxd and no Outlook.sqlite "
                                   "found"}, 2
        db_path = paths[0]
    return build_snapshot(db_path, days)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default=None)
    ap.add_argument("--days", type=int, default=None)
    ap.add_argument("--out", default="calendar.json")
    args = ap.parse_args()

    repo_dir = os.path.dirname(os.path.abspath(__file__))
    config = odb.load_config(repo_dir)
    days = args.days or int(config.get("window_days") or 14)

    snapshot, diagnostics, code = build_snapshot_auto(
        config, days, db_override=args.db)
    if snapshot is None:
        diag_path = os.path.join(os.path.dirname(
            os.path.abspath(args.out)), "export-diagnostics.json")
        with open(diag_path, "w") as fh:
            json.dump(diagnostics, fh, indent=2, default=str)
        print("ERROR: export failed (code %s). Diagnostics: %s"
              % (code, diag_path), file=sys.stderr)
        return code
    with open(args.out, "w") as fh:
        json.dump(snapshot, fh, indent=2)
    print("Wrote %s: %d events (source: %s)" % (
        args.out, len(snapshot["events"]), snapshot["source"]["app"]))
    if code == 4:
        print("WARNING: zero events decoded — snapshot written anyway so "
              "the failure is visible downstream. Stats: %s"
              % json.dumps(diagnostics.get("stats")), file=sys.stderr)
    return code


if __name__ == "__main__":
    sys.exit(main())
