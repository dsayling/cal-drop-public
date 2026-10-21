"""Fixture-based tests for cal-drop.

Builds throwaway SQLite databases in several ASSUMED Outlook for Mac
schemas and runs the real probe + exporter against them. Stdlib only:
    python3 -m unittest discover -s tests -v
"""

import json
import os
import sqlite3
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime, timedelta, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import export as exporter  # noqa: E402

NOW = datetime.now(timezone.utc)


def julian(dt):
    return dt.timestamp() / 86400.0 + 2440587.5


def cocoa(dt):
    return dt.timestamp() - 978307200


def at(day_offset, hour):
    base = (NOW + timedelta(days=day_offset)).replace(
        hour=hour, minute=0, second=0, microsecond=0)
    return base


def build_fixture_a(path):
    """Legacy-flavoured: CalendarEvents, Julian-day REAL dates, enums."""
    conn = sqlite3.connect(path)
    conn.execute("""
        CREATE TABLE CalendarEvents (
            EventID INTEGER PRIMARY KEY,
            EventSubject TEXT,
            EventStartDate REAL,
            EventEndDate REAL,
            EventAllDay INTEGER,
            EventRecurrence TEXT,
            EventSensitivity INTEGER,
            EventShowAs INTEGER,
            EventCancelled INTEGER
        )""")
    rows = [
        (1, "Ahsan sync", julian(at(1, 10)), julian(at(1, 11)), 0, None, 0, 2, 0),
        (2, "busy time", julian(at(0, 9)), julian(at(0, 10)), 0,
         "RRULE:FREQ=WEEKLY;COUNT=4", 0, 2, 0),
        (3, "Therapy payment details", julian(at(2, 12)),
         julian(at(2, 13)), 0, None, 2, 2, 0),
        (4, "Cancelled thing", julian(at(1, 15)), julian(at(1, 16)), 0,
         None, 0, 2, 1),
        (5, "Old meeting", julian(at(-30, 9)), julian(at(-30, 10)), 0,
         None, 0, 2, 0),
    ]
    conn.executemany("INSERT INTO CalendarEvents VALUES (?,?,?,?,?,?,?,?,?)",
                     rows)
    conn.execute("CREATE TABLE Folders (FolderID INTEGER, FolderName TEXT)")
    conn.commit()
    conn.close()


def build_fixture_b(path):
    """Text dates: Events(Title, Start, End, UID, AllDay)."""
    conn = sqlite3.connect(path)
    conn.execute("""
        CREATE TABLE Events (
            UID TEXT, Title TEXT, Start TEXT, End TEXT, AllDay INTEGER
        )""")
    fmt = "%Y-%m-%dT%H:%M:%S"
    conn.execute("INSERT INTO Events VALUES (?,?,?,?,?)", (
        "u1", "Design review",
        at(3, 14).strftime(fmt), at(3, 15).strftime(fmt), 0))
    conn.execute("INSERT INTO Events VALUES (?,?,?,?,?)", (
        "u2", "Conference day",
        at(5, 0).strftime("%Y-%m-%d"), at(6, 0).strftime("%Y-%m-%d"), 1))
    conn.commit()
    conn.close()


def build_fixture_c(path):
    """Cocoa-epoch numeric dates: Calendar(Event_Subject, ...)."""
    conn = sqlite3.connect(path)
    conn.execute("""
        CREATE TABLE Calendar (
            Event_Subject TEXT, Event_Start REAL, Event_End REAL
        )""")
    conn.execute("INSERT INTO Calendar VALUES (?,?,?)", (
        "Roadmap chat", cocoa(at(2, 9)), cocoa(at(2, 10))))
    conn.commit()
    conn.close()


class ExportTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="caldrop-test-")

    def snapshot(self, builder):
        db = os.path.join(self.tmp, "Outlook.sqlite")
        builder(db)
        snap, diag, code = exporter.build_snapshot(db, 14)
        self.assertIsNotNone(snap, "diagnostics: %s" % diag)
        return snap, diag, code

    def titles(self, snap):
        return [e["title"] for e in snap["events"]]

    def test_fixture_a_julian_enums_recurrence_private(self):
        snap, diag, code = self.snapshot(build_fixture_a)
        self.assertEqual(diag["table"], "CalendarEvents")
        self.assertEqual(diag["date_decoder"], "julian_day")
        titles = self.titles(snap)
        self.assertIn("Ahsan sync", titles)
        self.assertNotIn("Cancelled thing", titles)
        self.assertNotIn("Old meeting", titles)
        # weekly x4 starting today: today + 3 future occurrences in window
        self.assertEqual(titles.count("busy time"), 2)
        # sensitivity=2 redacted
        self.assertIn("Private appointment", titles)
        self.assertNotIn("Therapy payment details", titles)
        ahsan = [e for e in snap["events"] if e["title"] == "Ahsan sync"][0]
        self.assertEqual(ahsan["show_as"], "busy")
        self.assertFalse(ahsan["all_day"])
        self.assertEqual(code, 0)

    def test_fixture_b_text_dates(self):
        snap, diag, code = self.snapshot(build_fixture_b)
        self.assertEqual(diag["table"], "Events")
        titles = self.titles(snap)
        self.assertIn("Design review", titles)
        conf = [e for e in snap["events"]
                if e["title"] == "Conference day"][0]
        self.assertTrue(conf["all_day"])

    def test_fixture_c_cocoa_dates(self):
        snap, diag, code = self.snapshot(build_fixture_c)
        self.assertEqual(diag["date_decoder"], "cocoa_seconds")
        self.assertIn("Roadmap chat", self.titles(snap))


class AccessProbeTests(unittest.TestCase):
    def test_probe_names_what_the_process_can_see(self):
        probe = exporter.probe_store_access()
        self.assertIn("executable", probe)
        self.assertIn("exists", probe["profiles"])
        # One of the two outcomes must be recorded: entries seen, or the
        # exact error that stopped the listing.
        self.assertTrue("entries" in probe["profiles"] or
                        "list_error" in probe["profiles"])
        if "list_error" in probe["profiles"] or \
                not probe["profiles"]["exists"]:
            self.assertIn("hint", probe)


class ProbeTests(unittest.TestCase):
    def test_probe_report_has_schema_and_no_titles(self):
        tmp = tempfile.mkdtemp(prefix="caldrop-probe-")
        db = os.path.join(tmp, "Outlook.sqlite")
        build_fixture_a(db)
        out = os.path.join(tmp, "probe-report.json")
        proc = subprocess.run(
            [sys.executable, os.path.join(ROOT, "probe.py"),
             "--db", db, "--out", out],
            capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        with open(out) as fh:
            report_text = fh.read()
        self.assertIn("CalendarEvents", report_text)
        self.assertNotIn("Ahsan sync", report_text)
        self.assertNotIn("Therapy payment details", report_text)
        report = json.loads(report_text)
        self.assertEqual(report["databases"][0]["chosen_date_decoder"],
                         "julian_day")


def build_fixture_d(path):
    """The first real machine's shape: CalendarEvents is an index with
    start/end/UID and a PathToDataFile pointer but NO subject column and
    zero rows; Contacts has a display name and a start date and one row.
    v1 scored Contacts as the event table and would have exported a
    contact as an event."""
    conn = sqlite3.connect(path)
    conn.execute("""
        CREATE TABLE CalendarEvents (
            Record_RecordID INTEGER, PathToDataFile TEXT,
            Record_ModDate DATETIME, Calendar_EndDateUTC DATETIME,
            Calendar_StartDateUTC DATETIME, Calendar_UID TEXT,
            Record_UUID TEXT
        )""")
    conn.execute("""
        CREATE TABLE Contacts (
            Record_RecordID INTEGER, PathToDataFile TEXT,
            Contact_DisplayName TEXT, Record_StartDate DATETIME,
            Record_UUID TEXT
        )""")
    conn.execute("INSERT INTO Contacts VALUES (1, 'Data/x', 'Some Person', "
                 "2460000.5, 'u-1')")
    conn.commit()
    conn.close()


class RealShapeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="caldrop-test-")
        self.db = os.path.join(self.tmp, "Outlook.sqlite")
        build_fixture_d(self.db)

    def test_contacts_cannot_win(self):
        import outlook_db as odb
        conn = sqlite3.connect(self.db)
        conn.row_factory = sqlite3.Row
        cands = odb.table_candidates(odb.list_tables(conn))
        conn.close()
        by_name = {c["table"]: c for c in cands}
        self.assertTrue(by_name["Contacts"]["blocked_non_event"])
        self.assertLessEqual(by_name["Contacts"]["score"], 2)
        self.assertNotEqual(
            by_name["Contacts"]["columns"]["end"], "PathToDataFile")
        cal = by_name["CalendarEvents"]
        self.assertEqual(cal["columns"]["start"], "Calendar_StartDateUTC")
        self.assertEqual(cal["columns"]["end"], "Calendar_EndDateUTC")

    def test_export_refuses_instead_of_faking_an_event(self):
        snap, diag, code = exporter.build_snapshot(self.db, 14)
        self.assertIsNone(snap)
        self.assertEqual(code, 3)


class ProbeV2Tests(unittest.TestCase):
    def test_filesystem_map_counts_without_contents(self):
        tmp = tempfile.mkdtemp(prefix="caldrop-probe2-")
        root = os.path.join(tmp, "Outlook")
        data = os.path.join(root, "Outlook 15 Profiles", "Main Identity",
                            "Data", "Calendar")
        os.makedirs(data)
        secret = os.path.join(data, "12345.olk15Calendar")
        with open(secret, "w") as fh:
            fh.write("SECRET EVENT TITLE THAT MUST NOT LEAK")
        db = os.path.join(tmp, "Outlook.sqlite")
        build_fixture_a(db)
        out = os.path.join(tmp, "probe-report.json")
        proc = subprocess.run(
            [sys.executable, os.path.join(ROOT, "probe.py"),
             "--db", db, "--root", root, "--out", out],
            capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        with open(out) as fh:
            text = fh.read()
        self.assertNotIn("SECRET EVENT TITLE", text)
        report = json.loads(text)
        self.assertEqual(report["tool"], "cal-drop probe v2")
        fs = report["filesystem"][0]
        self.assertTrue(fs["exists"])
        self.assertEqual(fs["profiles"][0]["name"], "Main Identity")
        self.assertIn("Calendar", fs["profiles"][0]["data_entries"])
        self.assertEqual(fs["by_extension"][".olk15calendar"], 1)


class SyncRemoteTests(unittest.TestCase):
    def test_checkout_origin_wins_over_default(self):
        import sync as syncer
        tmp = tempfile.mkdtemp(prefix="caldrop-git-")
        subprocess.run(["git", "init", "-q"], cwd=tmp, check=True)
        subprocess.run(["git", "remote", "add", "origin",
                        "git@personal.github.com:your-username/cal-drop.git"],
                       cwd=tmp, check=True)
        config = {"remote": "git@github.com:your-username/cal-drop.git",
                  "remote_explicit": False}
        remote, source = syncer.resolve_remote(config, repo_dir=tmp)
        self.assertEqual(remote, "git@personal.github.com:your-username/cal-drop.git")
        self.assertEqual(source, "checkout origin")

    def test_explicit_config_wins_over_origin(self):
        import sync as syncer
        config = {"remote": "git@example.com:x/y.git",
                  "remote_explicit": True}
        remote, source = syncer.resolve_remote(config, repo_dir="/nonexistent")
        self.assertEqual(remote, "git@example.com:x/y.git")
        self.assertEqual(source, "config")


if __name__ == "__main__":
    unittest.main()
