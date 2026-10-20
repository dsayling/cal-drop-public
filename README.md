# cal-drop

A dumb calendar dead drop. A small Python job on your Mac reads the
**local Outlook for Mac cache** that Outlook already keeps in sync,
strips each event down to title + start + end + all-day + show-as,
and publishes the next 14 days as `calendar.json` on a `data`
branch in **your own private repo**. Any reader you control — a
script, an assistant, another machine — can then read that file,
which is useful when your calendar's normal sharing is free/busy
only and shows no titles.

No Microsoft API, no token, no app registration. Outlook does the
syncing; this just reads its local cache. Stdlib only — no pip
installs.

> **Warning: your instance must be private.**
> This public repo is only the code. Never run `sync.py` against
> a public repo: the `data` branch it pushes contains your real
> meeting titles. Set up your own private repo first (below).

## Privacy rules (enforced in export.py, before anything is written)

- Only five fields ever leave the Mac: `title`, `start`, `end`, `all_day`,
  `show_as`.
- No attendees, no bodies, no locations, no meeting links, no organizer.
- Events the database marks with sensitivity = private are exported with
  the title `Private appointment` on the SQLite path. The HxStore path
  does not map the sensitivity field yet — see ASSUMPTIONS.md.
- The `data` branch is a single orphan commit, force-pushed each sync, so
  the repo does not build a history of titles. (GitHub may retain
  unreachable commits for a while; that is a known limit, not a surprise.)
- The probe report contains schema only — no titles — so it is safe to
  share when something breaks.

## Layout

Two repos, kept separate:

- **This public repo** — the code: exporter, probe, sync, launchd
  agent, tests. No calendar data, ever.
- **Your private repo** — a copy of this code on its `main` branch,
  plus a `data` branch with one commit: `calendar.json` + a note.
  Read `generated_at` first; if it is old, the Mac was off or the
  sync is broken, and the events are stale.

Develop changes in the public copy, then pull them into your
private checkout (`git remote add upstream <public-repo-url>`,
`git pull upstream main`). Do not push from your private checkout
back to the public repo — that is the path that could leak data.

## Setup

1. Create a **private** GitHub repo for your instance, e.g.
   `cal-drop-private`.
2. Push this code to it as its `main` branch.
3. On the Mac running Outlook:

```bash
git clone git@github.com:YOUR-USERNAME/cal-drop-private.git cal-drop
cd cal-drop
./install.sh
```

Clone with whatever remote works on that Mac — an ssh-config host alias
is fine. `sync.py` publishes through the checkout's own `origin` (unless
`config.local.json` sets a `remote` explicitly), so the alias and its key
are the ones used for pushes too.

`install.sh` checks tools, installs the launchd agent (runs at login,
every 3 hours, and daily at 06:45), runs the probe, and does a first sync.

Optional local config: copy `config.example.json` to `config.local.json`
(gitignored) to override the database path, window size, or remote.

## Manual runs

```bash
python3 probe.py     # writes probe-report.json (schema only)
python3 hxprobe.py   # writes hxstore-report.json (HxStore spike, counts only)
python3 export.py    # writes ./calendar.json from the local cache
python3 sync.py      # export + push to the data branch if events changed
./run.sh             # same as sync.py, with launchd-friendly logging
```

The new Outlook (16.x) keeps its live data in `HxStore.hxd` (the Cola
store), not in the legacy `Outlook.sqlite`, which sits empty on current
installs. Export therefore reads HxStore first (`hxexport.py`:
appointment objects are recognised by their header signature, start/end
are the repeated adjacent .NET-ticks pair, the title follows the block
string order); the SQLite exporter remains as a fallback. `hxprobe.py`
is the counts-only spike that proved the format out, and `hxfind.py` is
the differential tool that mapped the fields against a synthetic test
event — its report is the one file that contains record contents, and
only for records matching the search string you give it.

Logs: `~/Library/Logs/cal-drop.log` and `.local/last-run.log`.
Diagnostics from a failed export: `.local/export-diagnostics.json`.

## Known gaps

See ASSUMPTIONS.md for the full list. The main ones:

- HxStore path: sensitivity (private) field and show-as field are not
  mapped yet; `show_as` is exported as `busy`, and recurring events are
  emitted as their stored occurrence only.
- All-day detection in `hxexport.py` is tuned for the US Mountain
  timezone offset (local midnight = 06:00/07:00 UTC). Other timezones
  may need their local-midnight UTC hours added.
- A few real-world records are not clean appointment objects; their
  exported "title" can be a line from the body instead of the subject.

## Fixing round

This was built on assumptions (see ASSUMPTIONS.md) and needed a fixing
round on the first real machine it ran on:

1. Run `python3 probe.py` and share `probe-report.json`.
2. The report shows the actual tables/columns and which date decoder fits;
   column candidates and decoders in `outlook_db.py` get adjusted to match.
3. Re-run `./run.sh`; the `source` block inside `calendar.json` records
   the table, columns, and decoder used, so a wrong guess is visible in
   the output itself.

## Tests

```bash
python3 -m unittest discover -s tests -v
```

The fixtures build throwaway SQLite databases in several assumed Outlook
schemas (Julian-day dates, ISO text dates, Cocoa-epoch dates) and
synthetic HxStore files, then run the real probe and exporters against
them, including recurrence expansion and private-title redaction.
