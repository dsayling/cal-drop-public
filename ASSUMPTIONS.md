# Assumptions

Built without access to the target machine, on purpose: every assumption
below is either detected at runtime or recorded in the output, so the
first real run tells us which ones were wrong.

1. **Where the data lives.** The legacy assumption — calendar in
   `Data/Outlook.sqlite` — is dead on the real machine: that database is
   an empty pre-migration index (0 mail/calendar rows, untouched for
   months). The new Outlook (16.x) keeps its live store in
   `HxStore.hxd` in the profile root, a page-based binary format
   (Microsoft's "Cola" engine: 4096-byte pages, LZ4-block-compressed
   records, .NET-ticks timestamps), reverse-engineered publicly in
   hxstore-decode's HXSTORE.md. Calendar items exist in it as
   Appointment (107) / Calendar (104) objects. `hxprobe.py` is the
   counts-only spike that verifies they decode on the real file before
   any exporter is built on this format. Appointment field layout was
   then mapped by differential analysis (a synthetic test event with
   known title/times, found with hxfind.py): an appointment record is a
   Cola object with uint32 1 at offset 0, its own length−4 at offset 4,
   and object type 107 (Appointment) as a uint16 at offset 29. Start
   and end are adjacent int64 .NET ticks, repeated in the record's
   second property block; the reminder is a lone tick before them.
   Strings follow a fixed block order — timezone names, a hex
   timezone-definition blob, title, organizer, location.
   `hxexport.py` implements this and export/sync prefer HxStore over
   the SQLite path. Known gaps in the HxStore exporter: the
   sensitivity (private) field is NOT mapped yet, so the SQLite path's
   "Private appointment" redaction has no equivalent here yet; show-as
   is exported as "busy" (the field is unmapped); recurring events are
   emitted as their stored occurrence only, with no expansion.
2. **The calendar schema is undocumented and version-dependent.** No table
   or column name is hardcoded as the only option. Tables and columns are
   scored by name (subject/title, start/end, recurrence, sensitivity...),
   and the winning mapping is written into `calendar.json` under `source`.
3. **Date storage encoding is unknown.** Candidates: ISO text, unix
   seconds/millis, Julian day, Cocoa/Core Data seconds (2001 epoch),
   Windows FILETIME, Excel days. The decoder is chosen by sampling the
   start column and picking the interpretation that lands the most values
   in a plausible range; the choice is recorded as `date_decoder`.
   Decoded dates are treated as local wall time converted via the Mac's
   timezone when the stored value has no zone of its own.
4. **Recurring events may be stored as master + rule text** (ICS RRULE),
   as master + an opaque blob, or as expanded rows. RRULE text is expanded
   in Python for the export window (daily/weekly/monthly/yearly, interval,
   count/until, byday/bymonthday). Opaque recurrence counts are reported
   in `source.stats.recurrence_unparsed` — if that number is high, the
   fixing round targets the blob format.
5. **Sensitivity value 2 means private** in the legacy enums. If the real
   schema uses another marker, private titles could leak into the export
   until the fixing round — check the first `calendar.json` by eye before
   trusting the sync. This is the one assumption with a privacy edge, and
   it is called out here rather than buried.
6. **Show-as enum** follows the legacy free/busy mapping
   (0 free, 1 tentative, 2 busy, 3 out of office). Wrong mapping only
   mislabels availability; titles and times are unaffected.
7. **The work Mac has python3 (Xcode CLT), git, and a passphrase-usable
   SSH key for GitHub in the launchd environment.** Some setups route
   personal repos through an ssh-config host alias (e.g.
   `git@personal.github.com`), so sync publishes through the checkout's
   own `origin` rather than a hardcoded `git@github.com` remote — the
   alias is what selects the right key.
   If the key needs an agent that launchd cannot reach, pushes fail with
   a clear log line and the fix is a dedicated key or
   `ssh-add --apple-use-keychain`.
8. **Outlook is running often enough** that its cache is fresh when the
   job runs (at login, every 3h, daily 06:45). `generated_at` in the
   snapshot is the staleness signal if it is not.

## What "broken" will probably look like, in order

- Probe finds the DB but scores no event table → column names differ;
  fix = extend the candidate lists in `outlook_db.py` using the report.
- Export runs, zero events → wrong decoder or the chosen table is a
  metadata table; fix = decoder choice / table score from diagnostics.
- Events appear but recurring blocks (busy time, lunch) show once →
  recurrence is a blob; fix = parse that format or expand from the
  master's rule columns the probe reveals.
- Everything works by hand, nothing arrives on schedule → launchd
  environment (PATH, SSH key, Mac asleep); fix in `run.sh` / the plist.
