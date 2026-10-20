#!/usr/bin/env python3
"""cal-drop probe v2: inspect the local Outlook for Mac storage and write
probe-report.json.

The report contains METADATA ONLY — table/column names, row counts, file
paths, sizes and dates. It never reads event titles, message subjects, or
file contents, so it is safe to share for the fixing round.

v1 answered "what does the one Outlook.sqlite we found look like?" On the
first real machine that database was an empty legacy index: 0 calendar
rows, and CalendarEvents had no title column at all — titles must live in
the per-item data files its PathToDataFile column points at. So v2 also
maps the filesystem: every Outlook profile, the shape of each profile's
Data folder (counts / sizes / newest files), every other database file
under the Outlook containers, and the Outlook app version.

Usage: python3 probe.py [--db PATH] [--out probe-report.json] [--no-fs]
       python3 probe.py --root DIR   (extra/replacement root; repeatable)
"""

import argparse
import json
import os
import plistlib
import shutil
import sys
from datetime import datetime, timezone

import outlook_db as odb

DB_EXTS = (".sqlite", ".db", ".sqlite3")
OTHER_DB_SIZE_LIMIT = 100 * 1024 * 1024


def _iso_mtime(path):
    try:
        return datetime.fromtimestamp(
            os.path.getmtime(path), timezone.utc).isoformat()
    except OSError:
        return None


def outlook_roots(home=None):
    home = home or os.path.expanduser("~")
    return [
        os.path.join(home, "Library", "Group Containers",
                     "UBF8T346G9.Office", "Outlook"),
        os.path.join(home, "Library", "Containers",
                     "com.microsoft.Outlook"),
        os.path.join(home, "Library", "Group Containers",
                     "UBF8T346G9.Office"),
    ]


def scan_tree(root):
    """Count/size/date summary of a directory tree. Paths are relative to
    root; contents are never opened."""
    by_ext = {}
    by_top = {}
    largest = []
    newest = []
    file_count = 0
    total_bytes = 0
    db_files = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort()
        for name in sorted(filenames):
            path = os.path.join(dirpath, name)
            try:
                st = os.stat(path)
            except OSError:
                continue
            rel = os.path.relpath(path, root)
            file_count += 1
            total_bytes += st.st_size
            ext = os.path.splitext(name)[1].lower() or "(none)"
            by_ext[ext] = by_ext.get(ext, 0) + 1
            top = rel.split(os.sep)[0] if os.sep in rel else "(root)"
            slot = by_top.setdefault(top, {"files": 0, "bytes": 0})
            slot["files"] += 1
            slot["bytes"] += st.st_size
            entry = {"path": rel, "size_bytes": st.st_size,
                     "mtime": datetime.fromtimestamp(
                         st.st_mtime, timezone.utc).isoformat()}
            largest.append(entry)
            newest.append(entry)
            if name.lower().endswith(DB_EXTS):
                db_files.append({"path": path, "size_bytes": st.st_size,
                                 "mtime": entry["mtime"]})
    largest.sort(key=lambda e: e["size_bytes"], reverse=True)
    newest.sort(key=lambda e: e["mtime"], reverse=True)
    return {
        "file_count": file_count,
        "total_bytes": total_bytes,
        "by_extension": dict(sorted(
            by_ext.items(), key=lambda kv: kv[1], reverse=True)[:20]),
        "by_top_entry": dict(sorted(
            by_top.items(), key=lambda kv: kv[1]["bytes"],
            reverse=True)[:20]),
        "largest_files": largest[:15],
        "newest_files": newest[:15],
        "database_files": db_files,
    }


def scan_filesystem(roots):
    out = []
    for root in roots:
        if not os.path.isdir(root):
            out.append({"root": root, "exists": False})
            continue
        entry = {"root": root, "exists": True}
        entry.update(scan_tree(root))
        profiles_dir = os.path.join(root, "Outlook 15 Profiles")
        if os.path.isdir(profiles_dir):
            profiles = []
            for name in sorted(os.listdir(profiles_dir)):
                pdir = os.path.join(profiles_dir, name)
                if not os.path.isdir(pdir):
                    continue
                prof = {"name": name}
                prof.update(scan_tree(pdir))
                data_dir = os.path.join(pdir, "Data")
                if os.path.isdir(data_dir):
                    prof["data_entries"] = sorted(os.listdir(data_dir))
                profiles.append(prof)
            entry["profiles"] = profiles
        out.append(entry)
    return out


def scan_other_databases(db_paths, known_paths):
    """Schema-only scan of database files that are not the primary DBs."""
    out = []
    for path in db_paths:
        if path in known_paths:
            continue
        entry = {"path": path, "size_bytes": None, "mtime": _iso_mtime(path)}
        try:
            entry["size_bytes"] = os.path.getsize(path)
            if entry["size_bytes"] > OTHER_DB_SIZE_LIMIT:
                entry["skipped"] = "over %d bytes; not opened" % \
                    OTHER_DB_SIZE_LIMIT
                out.append(entry)
                continue
            tmpdir, conn = odb.open_db_copy(path)
            try:
                tables = odb.list_tables(conn)
                entry["table_count"] = len(tables)
                entry["tables"] = [{
                    "name": t["name"], "row_count": t["row_count"],
                    "columns": t["columns"],
                } for t in tables]
                entry["candidate_event_tables"] = \
                    odb.table_candidates(tables)[:3]
            finally:
                conn.close()
                shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception as exc:
            entry["error"] = str(exc)
        out.append(entry)
    return out


def app_info():
    path = "/Applications/Microsoft Outlook.app/Contents/Info.plist"
    if not os.path.exists(path):
        return {"path": path, "exists": False}
    try:
        with open(path, "rb") as fh:
            plist = plistlib.load(fh)
        return {"path": path, "exists": True,
                "short_version": plist.get("CFBundleShortVersionString"),
                "bundle_version": plist.get("CFBundleVersion")}
    except Exception as exc:
        return {"path": path, "exists": True, "error": str(exc)}


def describe_database(path):
    entry = {"path": path}
    try:
        entry["size_bytes"] = os.path.getsize(path)
        entry["mtime"] = _iso_mtime(path)
        tmpdir, conn = odb.open_db_copy(path)
        try:
            tables = odb.list_tables(conn)
            cands = odb.table_candidates(tables)
            entry["table_count"] = len(tables)
            entry["tables"] = [{
                "name": t["name"], "row_count": t["row_count"],
                "columns": t["columns"],
            } for t in tables]
            entry["candidate_event_tables"] = cands[:5]
            top = cands[0] if cands else None
            if top and top["columns"]["start"]:
                col = top["columns"]["start"]
                try:
                    vals = [r[0] for r in conn.execute(
                        'SELECT "%s" FROM "%s" WHERE "%s" IS NOT NULL '
                        "LIMIT 500" % (col, top["table"], col)).fetchall()]
                    entry["start_column_value_types"] = sorted(
                        {type(v).__name__ for v in vals})
                    entry["chosen_date_decoder"] = odb.choose_decoder(vals)
                except Exception as exc:
                    entry["start_column_error"] = str(exc)
        finally:
            conn.close()
            shutil.rmtree(tmpdir, ignore_errors=True)
    except Exception as exc:
        entry["error"] = str(exc)
    return entry


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default=None,
                    help="explicit Outlook.sqlite path (default: discover)")
    ap.add_argument("--out", default="probe-report.json")
    ap.add_argument("--no-fs", action="store_true",
                    help="skip the filesystem map (schema only, like v1)")
    ap.add_argument("--root", action="append", default=None,
                    help="filesystem root to map instead of the defaults "
                         "(repeatable; used by tests)")
    args = ap.parse_args()

    config = odb.load_config(os.path.dirname(os.path.abspath(__file__)))
    db_path = args.db or config.get("db_path")
    paths = [db_path] if db_path else odb.find_db_paths()

    report = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "tool": "cal-drop probe v2",
        "note": "Metadata only: schema, row counts, file paths, sizes and "
                "dates. No event titles, subjects, bodies, or file "
                "contents are read.",
        "outlook_app": app_info(),
        "databases": [],
    }
    if not paths:
        report["error"] = (
            "No Outlook.sqlite found under the known profile locations. "
            "Either Outlook for Mac has never run under this user, or the "
            "profile lives somewhere else — that is fix-round information.")
    for path in paths:
        report["databases"].append(describe_database(path))

    if not args.no_fs:
        roots = args.root if args.root else outlook_roots()
        report["filesystem"] = scan_filesystem(roots)
        all_db_files = []
        for root_entry in report["filesystem"]:
            for dbf in root_entry.get("database_files", []):
                if dbf["path"] not in all_db_files:
                    all_db_files.append(dbf["path"])
        report["other_databases"] = scan_other_databases(
            all_db_files, set(paths))

    with open(args.out, "w") as fh:
        json.dump(report, fh, indent=2)
    print("Wrote %s" % args.out)
    for entry in report["databases"]:
        print("- %s" % entry["path"])
        if entry.get("error"):
            print("  ERROR: %s" % entry["error"])
        else:
            top = (entry.get("candidate_event_tables") or [{}])[0]
            print("  tables: %s, best event-table guess: %s (score %s), "
                  "date decoder: %s" % (
                      entry.get("table_count"), top.get("table"),
                      top.get("score"), entry.get("chosen_date_decoder")))
    for entry in report.get("filesystem", []):
        if entry.get("exists"):
            print("  fs %s: %d files, %d bytes, %d database files" % (
                entry["root"], entry.get("file_count", 0),
                entry.get("total_bytes", 0),
                len(entry.get("database_files", []))))
    for entry in report.get("other_databases", []):
        print("  other-db %s: %s" % (
            entry["path"],
            entry.get("error") or entry.get("skipped") or
            "%d tables" % entry.get("table_count", 0)))
    return 0 if paths else 1


if __name__ == "__main__":
    sys.exit(main())
