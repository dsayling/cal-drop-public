"""cal-drop exporter for the new Outlook's HxStore.hxd.

Field mapping established by differential analysis against a synthetic
test event with known title/times (2026-10-20; see ASSUMPTIONS.md):

- An appointment is a Cola object whose decompressed payload has
  uint32 1 at offset 0, its own length-4 at offset 4, and the
  Appointment object type (uint16 107) at offset 29.
- Start and end are adjacent int64 .NET ticks (start at n, end at
  n+8), repeated identically in the record's second property block.
  The reminder is a lone tick, so it never forms a false pair.
- Strings are length-prefixed UTF-16LE in a stable order per block:
  timezone names, a hex timezone-definition blob, then the title,
  the organizer, and the location (location is not exported).

Same output contract as the SQLite exporter: title, start, end,
all_day, show_as only — output by build_snapshot_hx has the same shape
as export.build_snapshot. Known gaps, recorded in ASSUMPTIONS.md: the
sensitivity (private) field and the show-as field are not mapped yet,
and recurring events are emitted as their stored occurrence only.
"""

import os
import re
import struct
from collections import Counter
from datetime import datetime, time, timedelta, timezone

import hxstore
import outlook_db as odb

APPT_OBJECT_TYPE = 107
SECONDS = hxstore.TICKS_PER_SECOND

_TZ_ID = re.compile(r"^[A-Za-z0-9_+-]+(/[A-Za-z0-9_+-]+)+$")
_TZ_POSIX = re.compile(r"^[A-Z]{2,5}[+-]?\d{1,2}([A-Z]{2,5}"
                       r"([+-]?\d{1,2})?)?$")
_TZ_PREFIX = re.compile(r"^([A-Za-z]+)/([A-Za-z_+-]+)")
_TZ_REGIONS = {"Africa", "America", "Antarctica", "Arctic", "Asia",
               "Atlantic", "Australia", "Brazil", "Canada", "Chile",
               "Etc", "Europe", "Indian", "Mexico", "Pacific", "US"}
_HEX = re.compile(r"^[0-9A-Fa-f]+$")
_BLOB = re.compile(r"^[A-Za-z0-9+/=_-]{20,}$")
_GUID = re.compile(r"^[0-9A-Fa-f]{8}-[0-9A-Fa-f-]{27,}$")


try:
    from zoneinfo import available_timezones
    _KNOWN_TZ = frozenset(available_timezones())
except Exception:
    _KNOWN_TZ = frozenset()


def _peel_tz_prefix(text):
    """Adjacent strings sometimes merge into one extracted run, e.g.
    'America/BahiaTeam 3.6 - Project Review...' — peel a leading IANA
    timezone id off and keep the remainder. The peel point is a
    lower->upper/digit boundary where the prefix is a real timezone
    id, so a title's own camel humps are never eaten."""
    for i in range(1, len(text)):
        if (text[i].isupper() or text[i].isdigit()) and \
                text[i - 1].islower():
            cand = text[:i]
            if "/" not in cand:
                continue
            if cand in _KNOWN_TZ or (
                    _TZ_ID.match(cand) and
                    cand.split("/")[0] in _TZ_REGIONS):
                return text[i:]
    return text


def is_appointment(payload):
    return len(payload) >= 64 and \
        struct.unpack_from("<I", payload, 0)[0] == 1 and \
        struct.unpack_from("<I", payload, 4)[0] == len(payload) - 4 and \
        struct.unpack_from("<H", payload, 29)[0] == APPT_OBJECT_TYPE


def find_start_end(payload):
    """The (start_ticks, end_ticks) pair: adjacent int64s, end after
    start, whole-second values, duration 1 minute to 14 days. When a
    pair repeats (the record's second block), it wins."""
    ticks = dict(hxstore.all_plausible_ticks(payload))
    counts = Counter()
    first_at = {}
    for off in sorted(ticks):
        if off + 8 not in ticks:
            continue
        start, end = ticks[off], ticks[off + 8]
        dur = (end - start) // SECONDS
        if end <= start or dur < 60 or dur > 14 * 86400:
            continue
        if start % SECONDS or end % SECONDS:
            continue
        counts[(start, end)] += 1
        first_at.setdefault((start, end), off)
    if not counts:
        return None
    return max(counts, key=lambda p: (counts[p], -first_at[p]))


def _is_tz(text):
    return bool(_TZ_ID.match(text)) or bool(_TZ_POSIX.match(text)) or \
        text.startswith("tzone://") or text.endswith(" Time") or \
        text in ("UTC", "GMT", "Zulu")


def _is_hex_blob(text):
    return len(text) >= 32 and bool(_HEX.match(text))


def extract_title(payload):
    """Title per the block string order: the first content string after
    the hex timezone blob. Fallbacks: first content string that repeats
    in the record, then the first content string."""
    strings = [_peel_tz_prefix(text.strip()) for _off, text in
               hxstore.utf16_strings_with_offsets(payload)]
    strings = [s for s in strings if s]

    def content(s):
        return not _is_tz(s) and not _is_hex_blob(s) and "@" not in s \
            and not _GUID.match(s) and not _BLOB.match(s) \
            and not s.startswith("-----")

    blob_idx = next((i for i, s in enumerate(strings)
                     if _is_hex_blob(s)), None)
    if blob_idx is not None:
        for s in strings[blob_idx + 1:]:
            if content(s):
                return s
    counts = Counter(s for s in strings if content(s))
    repeated = [s for s, n in counts.items() if n >= 2]
    if repeated:
        return repeated[0]
    for s in strings:
        if content(s):
            return s
    return None


def _is_all_day(start_ticks, end_ticks):
    dur = (end_ticks - start_ticks) // SECONDS
    if dur <= 0 or dur % 86400:
        return False
    tod = (start_ticks // SECONDS) % 86400
    # Local midnight seen from UTC: 00:00 (UTC itself) or 06:00 / 07:00
    # (for example US Mountain daylight / standard time).
    # NOTE: this check is tuned for the US Mountain offset; other
    # timezones may need their local-midnight UTC hours added here.
    return tod in (0, 6 * 3600, 7 * 3600)


def build_snapshot_hx(hx_path, days):
    """Mirror of export.build_snapshot for HxStore sources.
    Returns (snapshot|None, diagnostics, code)."""
    diagnostics = {"hx_path": hx_path}
    with open(hx_path, "rb") as fh:
        data = fh.read()
    primary, _sec, _counts = hxstore.discover_store_ids(data)
    if primary is None:
        diagnostics["error"] = "no data-page store id found"
        return None, diagnostics, 3

    best = {}
    scanned = 0
    gen = hxstore.iter_records(data, primary)
    while True:
        try:
            record_id, payload, _sa, _sb, _off, _fmt, _unk = next(gen)
        except StopIteration:
            break
        scanned += 1
        if record_id not in best or len(payload) > len(best[record_id]):
            best[record_id] = payload
    diagnostics["records_scanned"] = scanned
    diagnostics["unique_record_ids"] = len(best)

    stats = {"signature_matches": 0, "skipped_no_pair": 0,
             "skipped_no_title": 0, "out_of_window": 0,
             "duplicates_collapsed": 0}
    now_local = datetime.now(odb.LOCAL_TZ)
    window_start_local = datetime.combine(
        now_local.date(), time.min, tzinfo=odb.LOCAL_TZ)
    window_start = window_start_local.astimezone(timezone.utc)
    window_end = window_start + timedelta(days=days)

    events = []
    seen = set()
    for payload in best.values():
        if not is_appointment(payload):
            continue
        stats["signature_matches"] += 1
        pair = find_start_end(payload)
        if pair is None:
            stats["skipped_no_pair"] += 1
            continue
        start = hxstore.ticks_to_datetime(pair[0])
        end = hxstore.ticks_to_datetime(pair[1])
        if not (start < window_end and end > window_start):
            stats["out_of_window"] += 1
            continue
        title = extract_title(payload) or "(no title)"
        if title == "(no title)":
            stats["skipped_no_title"] += 1
        key = (title, pair[0], pair[1])
        if key in seen:
            stats["duplicates_collapsed"] += 1
            continue
        seen.add(key)
        events.append({
            "title": title,
            "start": start.astimezone(odb.LOCAL_TZ).isoformat(),
            "end": end.astimezone(odb.LOCAL_TZ).isoformat(),
            "all_day": _is_all_day(pair[0], pair[1]),
            "show_as": "busy",
        })
    events.sort(key=lambda e: (e["start"], e["title"]))
    diagnostics["stats"] = stats
    diagnostics["events_out"] = len(events)
    if not stats["signature_matches"]:
        diagnostics["error"] = "no appointment-signature records found"
        return None, diagnostics, 3
    snapshot = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "timezone": str(odb.LOCAL_TZ),
        "window_days": days,
        "source": {
            "app": "outlook-for-mac-hxstore",
            "profile_dir": os.path.basename(os.path.dirname(hx_path)),
            "stats": stats,
        },
        "events": events,
    }
    return snapshot, diagnostics, 0 if events else 4
