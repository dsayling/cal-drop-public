#!/usr/bin/env python3
"""cal-drop sync: export the local Outlook cache and, if the events changed,
publish calendar.json as a single-commit snapshot on the repo's data branch.

The data branch is always exactly one orphan commit (force-pushed), so the
repo never accumulates a history of meeting titles. Publishing works in a
throwaway git repo under .local/, using the work Mac's existing SSH key —
the same key that clones the code checkout.

Usage: python3 sync.py [--force]
Exit codes mirror export.py when the export itself fails; nothing is
pushed in that case, so a stale generated_at downstream IS the failure
signal.
"""

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime, timezone

REPO_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, REPO_DIR)

import export as exporter  # noqa: E402
import outlook_db as odb  # noqa: E402

DATA_README = """# cal-drop data branch

Machine-generated snapshot. One commit only, force-pushed by sync.py on
the work Mac. `calendar.json` holds title, start, end, all-day and
show-as for the next window of days — nothing else is exported.
Read `generated_at` first: if it is old, the sync is broken or the Mac
was off, and the events below are stale.
"""


def checkout_origin(repo_dir):
    """The origin URL of the code checkout itself, or None."""
    proc = subprocess.run(
        ["git", "-C", repo_dir, "remote", "get-url", "origin"],
        capture_output=True, text=True)
    if proc.returncode != 0:
        return None
    return proc.stdout.strip() or None


def resolve_remote(config, repo_dir=REPO_DIR):
    """Return (remote, source). An explicit remote in config.local.json /
    config.json wins; otherwise inherit the checkout's origin — that is
    the URL that actually cloned the repo on this machine, including any
    ssh-config host alias (e.g. git@personal.github.com) and its key. The
    built-in git@github.com default is the last resort."""
    if config.get("remote_explicit") and config.get("remote"):
        return config["remote"], "config"
    origin = checkout_origin(repo_dir)
    if origin:
        return origin, "checkout origin"
    return config["remote"], "default"


def run(cmd, cwd=None, check=True):
    proc = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if check and proc.returncode != 0:
        raise RuntimeError("command failed: %s\n%s%s" % (
            " ".join(cmd), proc.stdout, proc.stderr))
    return proc


def events_fingerprint(snapshot):
    payload = json.dumps(snapshot.get("events", []), sort_keys=True)
    return hashlib.sha256(payload.encode()).hexdigest()


def fetch_current_fingerprint(config):
    """Best-effort fingerprint of the calendar.json on the data branch."""
    tmp = tempfile.mkdtemp(prefix="caldrop-read-")
    try:
        run(["git", "init", "-q"], cwd=tmp)
        run(["git", "remote", "add", "origin", config["remote"]], cwd=tmp)
        proc = run(["git", "fetch", "-q", "--depth", "1", "origin",
                    config["data_branch"]], cwd=tmp, check=False)
        if proc.returncode != 0:
            return None
        proc = run(["git", "show", "FETCH_HEAD:calendar.json"], cwd=tmp,
                   check=False)
        if proc.returncode != 0:
            return None
        try:
            return events_fingerprint(json.loads(proc.stdout))
        except ValueError:
            return None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def publish(snapshot_path, config):
    tmp = tempfile.mkdtemp(prefix="caldrop-push-")
    try:
        run(["git", "init", "-q", "-b", "snapshot"], cwd=tmp)
        run(["git", "config", "user.name", config["git_name"]], cwd=tmp)
        run(["git", "config", "user.email", config["git_email"]], cwd=tmp)
        shutil.copy2(snapshot_path, os.path.join(tmp, "calendar.json"))
        with open(os.path.join(tmp, "README.md"), "w") as fh:
            fh.write(DATA_README)
        run(["git", "add", "calendar.json", "README.md"], cwd=tmp)
        run(["git", "commit", "-q", "-m", "calendar snapshot %s"
             % datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%MZ")],
            cwd=tmp)
        run(["git", "remote", "add", "origin", config["remote"]], cwd=tmp)
        run(["git", "push", "-q", "--force", "origin",
             "snapshot:%s" % config["data_branch"]], cwd=tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    force = "--force" in sys.argv
    config = odb.load_config(REPO_DIR)
    local_dir = os.path.join(REPO_DIR, ".local")
    os.makedirs(local_dir, exist_ok=True)
    snapshot_path = os.path.join(local_dir, "calendar.json")
    log_path = os.path.join(local_dir, "last-run.log")

    def log(msg):
        line = "%s %s" % (datetime.now(timezone.utc).isoformat(), msg)
        print(line)
        with open(log_path, "a") as fh:
            fh.write(line + "\n")

    remote, remote_source = resolve_remote(config)
    config["remote"] = remote
    log("Remote (%s): %s" % (remote_source, remote))

    snapshot, diagnostics, code = exporter.build_snapshot_auto(
        config, int(config.get("window_days") or 14))
    if snapshot is None:
        diag_path = os.path.join(local_dir, "export-diagnostics.json")
        with open(diag_path, "w") as fh:
            json.dump(diagnostics, fh, indent=2, default=str)
        log("ERROR: export failed (code %s). Diagnostics: %s"
            % (code, diag_path))
        return code
    with open(snapshot_path, "w") as fh:
        json.dump(snapshot, fh, indent=2)
    # A success clears the last failure's diagnostics — otherwise the
    # stale file keeps "reporting" an old error after the fix lands.
    stale_diag = os.path.join(local_dir, "export-diagnostics.json")
    if os.path.exists(stale_diag):
        os.remove(stale_diag)

    new_fp = events_fingerprint(snapshot)
    if not force and fetch_current_fingerprint(config) == new_fp:
        log("No change (%d events); nothing pushed."
            % len(snapshot["events"]))
        return 0
    try:
        publish(snapshot_path, config)
    except RuntimeError as exc:
        log("ERROR: push failed (SSH key / network / repo access?). "
            "Details in .local/last-run.log\n%s" % exc)
        return 5
    log("Pushed snapshot: %d events, generated_at=%s"
        % (len(snapshot["events"]), snapshot["generated_at"]))
    return 0 if code == 0 else code


if __name__ == "__main__":
    sys.exit(main())
