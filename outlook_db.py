"""Shared Outlook-for-Mac local-database reading logic for cal-drop.

Stdlib only. Everything in here is built on ASSUMPTIONS (see ASSUMPTIONS.md):
the Outlook for Mac profile store is a SQLite database, its calendar schema
is undocumented and version-dependent, and date storage has used several
encodings over the years. The strategy is runtime discovery, not fixed
queries: find candidate event tables/columns by name scoring, decide the
date encoding from the actual values, and report every choice made in a
diagnostics dict so a broken run can be fixed from its own output.
"""

import glob
import json
import os
import re
import shutil
import sqlite3
import tempfile
from datetime import datetime, date, time, timedelta, timezone

LOCAL_TZ = datetime.now().astimezone().tzinfo

DB_GLOB_PATTERNS = [
    "~/Library/Group Containers/UBF8T346G9.Office/Outlook/"
    "Outlook 15 Profiles/*/Data/Outlook.sqlite",
    "~/Library/Group Containers/UBF8T346G9.Office/Outlook/"
    "Outlook 15 Profiles/*/Outlook.sqlite",
    "~/Library/Containers/com.microsoft.Outlook/Data/Library/"
    "Application Support/Microsoft/Outlook/*/Data/Outlook.sqlite",
]

EXCLUDE_COL_SUBSTR = (
    "timezone", "tz", "offset", "reminder", "type_id", "folder",
)


def find_db_paths():
    found = []
    for pattern in DB_GLOB_PATTERNS:
        found.extend(glob.glob(os.path.expanduser(pattern)))
    found = sorted(set(found))
    found.sort(key=lambda p: os.path.getmtime(p) if os.path.exists(p) else 0,
               reverse=True)
    return found


def copy_db(db_path):
    """Copy the DB (plus -wal/-shm) to a temp dir; return (tmpdir, copy_path).

    Copying avoids locking fights with a running Outlook and avoids opening
    the live store at all.
    """
    tmpdir = tempfile.mkdtemp(prefix="caldrop-")
    dst = os.path.join(tmpdir, "Outlook.sqlite")
    shutil.copy2(db_path, dst)
    for suffix in ("-wal", "-shm"):
        src = db_path + suffix
        if os.path.exists(src):
            shutil.copy2(src, dst + suffix)
    return tmpdir, dst


def open_db_copy(db_path):
    tmpdir, copy_path = copy_db(db_path)
    conn = sqlite3.connect("file:%s?mode=ro" % copy_path, uri=True)
    conn.row_factory = sqlite3.Row
    return tmpdir, conn


def list_tables(conn):
    rows = conn.execute(
        "SELECT name FROM sqlite_master WHERE type IN ('table','view') "
        "AND name NOT LIKE 'sqlite_%' ORDER BY name").fetchall()
    tables = []
    for row in rows:
        name = row["name"]
        try:
            cols = [dict(r) for r in conn.execute(
                "PRAGMA table_info(%s)" % _quote_ident(name)).fetchall()]
            cols = [{"name": c["name"], "type": c["type"]} for c in cols]
        except sqlite3.Error:
            cols = []
        try:
            count = conn.execute(
                "SELECT COUNT(*) AS n FROM %s" % _quote_ident(name)
            ).fetchone()["n"]
        except sqlite3.Error:
            count = None
        tables.append({"name": name, "columns": cols, "row_count": count})
    return tables


def _quote_ident(name):
    return '"%s"' % name.replace('"', '""')


def _col_score(col_name, wanted_groups):
    """Score a column name against groups of substrings, best group first."""
    low = col_name.lower()
    if any(x in low for x in EXCLUDE_COL_SUBSTR):
        return -5
    for idx, group in enumerate(wanted_groups):
        if any(g in low for g in group):
            return 10 - idx
    return 0


SUBJECT_GROUPS = [["subject"], ["summary"], ["title"], ["eventname"], ["name"]]
START_GROUPS = [["startdate", "start_date", "startdatetime"],
                ["start"], ["begin", "from"]]
END_GROUPS = [["enddate", "end_date", "enddatetime"], ["end"], ["until"]]
ALLDAY_GROUPS = [["allday", "all_day"], ["isallday"]]
UID_GROUPS = [["uid"], ["guid"], ["globalobjid"], ["entryid"], ["recordid"],
              ["id"]]
RECURRENCE_GROUPS = [["recurrencepattern", "recurrencerule", "rrule"],
                     ["recurrence"], ["repeat"]]
SENSITIVITY_GROUPS = [["sensitivity"], ["private"], ["privacy"]]
SHOWAS_GROUPS = [["showas", "show_as"], ["legacyfreebusy", "freebusy"],
                 ["busystatus"], ["status"], ["availability"]]
CANCELLED_GROUPS = [["cancelled", "canceled"], ["ismeetingcanceled"]]


def _best_col(columns, groups, require_positive=True):
    best, best_score = None, 0
    for col in columns:
        s = _col_score(col, groups)
        if s > best_score:
            best, best_score = col, s
    if best is None and not require_positive:
        return None
    return best


NON_EVENT_TABLES = {"contacts", "tasks", "mail", "notes", "folders",
                    "categories", "rules", "settings", "signatures",
                    "conversations", "threads", "files", "main"}


def table_candidates(tables):
    """Return scored candidate event tables, best first."""
    out = []
    for t in tables:
        cols = [c["name"] for c in t["columns"]]
        low = t["name"].lower()
        base = low.split("_")[0]
        blocked = low in NON_EVENT_TABLES or base in NON_EVENT_TABLES
        calendarish = any(k in low for k in
                          ("calendar", "event", "appointment", "meeting"))
        score = 0
        if "calendar" in low:
            score += 3
        if "event" in low:
            score += 2
        if "appointment" in low or "meeting" in low:
            score += 2
        subject = _best_col(cols, SUBJECT_GROUPS)
        start = _best_col(cols, START_GROUPS)
        end = _best_col(cols, END_GROUPS)
        if subject:
            score += 3
        if start:
            score += 3
        if end:
            score += 2
        if t["row_count"]:
            score += 1
        if not (subject and start):
            score = min(score, 2)
        if blocked or not calendarish:
            # A contacts/tasks table with a name + a start-ish column is
            # not a calendar, however well its columns score. Real
            # CalendarEvents-with-no-subject stays visible in the list
            # (capped, with its columns) but can never be exported.
            score = min(score, 4 if calendarish else 2)
        out.append({
            "table": t["name"], "score": score,
            "blocked_non_event": blocked,
            "row_count": t["row_count"],
            "columns": {
                "subject": subject, "start": start, "end": end,
                "all_day": _best_col(cols, ALLDAY_GROUPS),
                "uid": _best_col(cols, UID_GROUPS),
                "recurrence": _best_col(cols, RECURRENCE_GROUPS),
                "sensitivity": _best_col(cols, SENSITIVITY_GROUPS),
                "show_as": _best_col(cols, SHOWAS_GROUPS),
                "cancelled": _best_col(cols, CANCELLED_GROUPS),
            },
            "all_columns": cols,
        })
    out.sort(key=lambda c: c["score"], reverse=True)
    return out


# ---------------------------------------------------------------- dates

UNIX_EPOCH = datetime(1970, 1, 1, tzinfo=timezone.utc)
COCOA_OFFSET = 978307200  # seconds between 1970-01-01 and 2001-01-01
FILETIME_OFFSET = 11644473600  # seconds between 1601-01-01 and 1970-01-01
PLAUSIBLE_START = datetime(2023, 1, 1, tzinfo=timezone.utc)
PLAUSIBLE_END = datetime(2028, 1, 1, tzinfo=timezone.utc)

TEXT_FORMATS = [
    "%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S.%f",
    "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M", "%Y-%m-%d", "%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M",
    "%d %b %Y %H:%M:%S", "%Y/%m/%d %H:%M:%S",
]


def _parse_text_date(value):
    text = value.strip()
    if not text:
        return None
    norm = text
    if norm.endswith("Z"):
        norm = norm[:-1] + "+0000"
    norm = re.sub(r"([+-]\d{2}):(\d{2})$", r"\1\2", norm)
    for fmt in TEXT_FORMATS:
        try:
            dt = datetime.strptime(norm, fmt)
        except ValueError:
            continue
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=LOCAL_TZ)
        return dt.astimezone(timezone.utc)
    try:
        dt = datetime.fromisoformat(text.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=LOCAL_TZ)
        return dt.astimezone(timezone.utc)
    except ValueError:
        return None


def _numeric_candidates(value):
    """Yield (decoder_name, datetime) interpretations of a numeric date."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return
    if 2_000_000 < v < 3_000_000:  # Julian day
        ts = (v - 2440587.5) * 86400
        yield "julian_day", UNIX_EPOCH + timedelta(seconds=ts)
    if 946_684_800 < v < 4_102_444_800:  # unix seconds
        yield "unix_seconds", UNIX_EPOCH + timedelta(seconds=v)
    if 946_684_800_000 < v < 4_102_444_800_000:  # unix milliseconds
        yield "unix_millis", UNIX_EPOCH + timedelta(milliseconds=v)
    if 0 < v < (4_102_444_800 - COCOA_OFFSET):  # Cocoa / Core Data
        yield "cocoa_seconds", UNIX_EPOCH + timedelta(
            seconds=v + COCOA_OFFSET)
    if v > 100_000_000_000_000_000:  # Windows FILETIME (100ns since 1601)
        yield "filetime", UNIX_EPOCH + timedelta(
            seconds=v / 10_000_000 - FILETIME_OFFSET)
    if 20_000 < v < 60_000:  # Excel-style days since 1899-12-30
        yield "excel_days", datetime(1899, 12, 30, tzinfo=timezone.utc) + \
            timedelta(days=v)


def decode_date(value, decoder=None):
    """Decode one raw date value to an aware UTC datetime, or None."""
    if value is None:
        return None
    if isinstance(value, datetime):
        dt = value
        return dt.astimezone(timezone.utc) if dt.tzinfo else \
            dt.replace(tzinfo=LOCAL_TZ).astimezone(timezone.utc)
    if isinstance(value, date):
        return datetime.combine(value, time.min, tzinfo=LOCAL_TZ) \
            .astimezone(timezone.utc)
    if isinstance(value, bytes):
        try:
            value = value.decode("utf-8", "strict")
        except UnicodeDecodeError:
            return None
    if isinstance(value, str):
        return _parse_text_date(value)
    if isinstance(value, (int, float)):
        for name, dt in _numeric_candidates(value):
            if decoder is None or name == decoder:
                if PLAUSIBLE_START <= dt <= PLAUSIBLE_END:
                    return dt
                if decoder is not None:
                    return dt
        return None
    return None


def choose_decoder(values):
    """Pick the numeric decoder that puts the most values in a sane range."""
    numeric = [v for v in values if isinstance(v, (int, float))]
    if not numeric:
        return "text_or_native"
    names = ["julian_day", "unix_seconds", "unix_millis", "cocoa_seconds",
             "filetime", "excel_days"]
    best_name, best_count = None, -1
    for name in names:
        count = 0
        for v in numeric[:500]:
            for cand_name, dt in _numeric_candidates(v):
                if cand_name == name and \
                        PLAUSIBLE_START <= dt <= PLAUSIBLE_END:
                    count += 1
                    break
        if count > best_count:
            best_name, best_count = name, count
    return best_name or "unknown"


# ---------------------------------------------------------------- rrule

WEEKDAYS = {"MO": 0, "TU": 1, "WE": 2, "TH": 3, "FR": 4, "SA": 5, "SU": 6}


def parse_rrule(text):
    """Extract an RRULE dict from ICS-style text, or None."""
    if not text:
        return None
    if isinstance(text, bytes):
        try:
            text = text.decode("utf-8", "ignore")
        except Exception:
            return None
    match = re.search(r"RRULE:([^\r\n]+)", text)
    if not match:
        match = re.search(r"(FREQ=[^\r\n;]+(?:;[^\r\n]+)?)", text)
        if not match:
            return None
        body = match.group(1)
    else:
        body = match.group(1)
    rule = {}
    for part in body.split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            rule[k.strip().upper()] = v.strip()
    return rule if "FREQ" in rule else None


def expand_rrule(dtstart, rule, window_start, window_end, cap=1000):
    """Yield occurrence datetimes of an RRULE within the window (UTC)."""
    freq = rule.get("FREQ", "")
    interval = int(rule.get("INTERVAL", "1") or 1)
    count_limit = int(rule["COUNT"]) if rule.get("COUNT") else None
    until = None
    if rule.get("UNTIL"):
        until = _parse_text_date(rule["UNTIL"])
        if until is None and re.fullmatch(r"\d{8}", rule["UNTIL"]):
            until = datetime.strptime(rule["UNTIL"], "%Y%m%d") \
                .replace(tzinfo=timezone.utc) + timedelta(days=1)
    byday = []
    for tok in (rule.get("BYDAY") or rule.get("BYWEEKDAY") or "").split(","):
        tok = tok.strip()
        if not tok:
            continue
        m = re.fullmatch(r"([+-]?\d+)?(MO|TU|WE|TH|FR|SA|SU)", tok)
        if m:
            byday.append((int(m.group(1)) if m.group(1) else None,
                          WEEKDAYS[m.group(2)]))
    bymonthday = [int(x) for x in
                  (rule.get("BYMONTHDAY") or "").split(",") if x.strip()]
    bymonth = [int(x) for x in
               (rule.get("BYMONTH") or "").split(",") if x.strip()]

    def in_window(dt):
        return window_start <= dt < window_end

    produced = 0
    made = 0
    cursor = dtstart

    def emit(dt):
        nonlocal produced, made
        produced += 1
        if count_limit and produced > count_limit:
            return False
        if until and dt > until:
            return False
        if dt >= window_end:
            return False
        if in_window(dt):
            made += 1
        return True

    results = []

    def consider(dt):
        if in_window(dt):
            results.append(dt)
        return emit(dt)

    if freq == "DAILY":
        while cursor < window_end and made + produced < cap + 1000:
            if not consider(cursor):
                break
            cursor += timedelta(days=interval)
            if produced > cap:
                break
    elif freq == "WEEKLY":
        days = sorted({d for _, d in byday}) or [dtstart.weekday()]
        week_start = cursor - timedelta(days=cursor.weekday())
        cur = week_start
        while cur < window_end:
            for d in days:
                dt = cur + timedelta(days=d, hours=dtstart.hour,
                                     minutes=dtstart.minute,
                                     seconds=dtstart.second)
                dt = dt.replace(tzinfo=dtstart.tzinfo)
                if dt >= dtstart:
                    if not consider(dt):
                        return results
            cur += timedelta(weeks=interval)
            if produced > cap:
                break
    elif freq == "MONTHLY":
        cur = dtstart
        while cur < window_end:
            cands = []
            if bymonthday:
                for md in bymonthday:
                    try:
                        cands.append(cur.replace(day=md))
                    except ValueError:
                        pass
            elif byday:
                for ordinal, wd in byday:
                    cands.append(_nth_weekday(cur, wd, ordinal or 1))
            else:
                cands.append(cur)
            for dt in sorted(c for c in cands if c):
                if dt >= dtstart:
                    if not consider(dt):
                        return results
            cur = _add_months(cur, interval)
            if produced > cap:
                break
    elif freq == "YEARLY":
        cur = dtstart
        while cur < window_end:
            months = bymonth or [dtstart.month]
            for mo in months:
                try:
                    dt = cur.replace(month=mo)
                except ValueError:
                    continue
                if dt >= dtstart:
                    if not consider(dt):
                        return results
            cur = _add_months(cur, 12 * interval)
            if produced > cap:
                break
    else:
        if in_window(dtstart):
            results.append(dtstart)
    return results


def _add_months(dt, months):
    month = dt.month - 1 + months
    year = dt.year + month // 12
    month = month % 12 + 1
    day = min(dt.day, [31, 29 if year % 4 == 0 and
                      (year % 100 != 0 or year % 400 == 0) else 28,
                      31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1])
    return dt.replace(year=year, month=month, day=day)


def _nth_weekday(dt, weekday, ordinal):
    first = dt.replace(day=1)
    shift = (weekday - first.weekday()) % 7
    if ordinal > 0:
        day = 1 + shift + (ordinal - 1) * 7
    else:
        # last / nth-from-last weekday of the month
        nxt = _add_months(first, 1)
        last = nxt - timedelta(days=1)
        back = (last.weekday() - weekday) % 7
        day = last.day - back + (ordinal + 1) * 7
    try:
        return dt.replace(day=day)
    except ValueError:
        return None


def is_truthy(value):
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, bytes):
        return value not in (b"", b"\x00")
    return str(value).strip().lower() in ("1", "true", "yes", "y", "t")


def is_private(row_value, col_name):
    if row_value is None:
        return False
    if isinstance(row_value, (int, float)):
        # Outlook legacy free/busy + sensitivity enums: 2 is commonly
        # "private" for sensitivity. Flagged as an assumption.
        return row_value == 2 and "sensitiv" in (col_name or "").lower()
    return str(row_value).strip().lower() in ("private", "2")


SHOW_AS_MAP = {0: "free", 1: "tentative", 2: "busy", 3: "out_of_office",
               4: "working_elsewhere"}


def map_show_as(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return SHOW_AS_MAP.get(int(value), str(int(value)))
    text = str(value).strip().lower()
    return text or None


def load_config(repo_dir):
    defaults = {
        "db_path": None,
        "window_days": 14,
        "remote": "git@github.com:YOUR-USERNAME/cal-drop-private.git",
        "data_branch": "data",
        "git_name": "cal-drop",
        "git_email": "cal-drop@localhost",
        "remote_explicit": False,
    }
    for name in ("config.local.json", "config.json"):
        path = os.path.join(repo_dir, name)
        if os.path.exists(path):
            with open(path) as fh:
                loaded = json.load(fh)
            if "remote" in loaded:
                defaults["remote_explicit"] = True
            defaults.update(loaded)
            break
    return defaults
