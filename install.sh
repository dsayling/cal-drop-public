#!/bin/bash
# cal-drop one-time setup on the work Mac.
#   1. sanity-checks python3 / git / an SSH key
#   2. installs the launchd agent (daily 06:45 + every 3h + at login)
#   3. runs the probe — probe-report.json is the file to share if the
#      first export is wrong or empty
#   4. runs one sync so the data branch gets its first real snapshot
set -u
REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$REPO_DIR"

echo "== cal-drop install =="
command -v python3 >/dev/null || { echo "python3 not found — install Xcode Command Line Tools: xcode-select --install"; exit 1; }
command -v git >/dev/null || { echo "git not found — install Xcode Command Line Tools: xcode-select --install"; exit 1; }
python3 --version
git --version
if [ ! -e "$HOME/.ssh/id_ed25519" ] && [ ! -e "$HOME/.ssh/id_rsa" ] && [ -z "${SSH_AUTH_SOCK:-}" ]; then
  echo "WARNING: no default SSH key found in ~/.ssh — pushes will fail until the GitHub key is in place or loaded in the agent."
fi

chmod +x run.sh sync.py export.py probe.py hxprobe.py hxfind.py

PLIST_DST="$HOME/Library/LaunchAgents/com.caldrop.plist"
mkdir -p "$HOME/Library/LaunchAgents"
sed -e "s|__REPO_DIR__|$REPO_DIR|g" -e "s|__HOME__|$HOME|g" \
  launchd/com.caldrop.plist > "$PLIST_DST"
echo "Installed launch agent: $PLIST_DST"
launchctl bootout "gui/$(id -u)" "$PLIST_DST" 2>/dev/null
if launchctl bootstrap "gui/$(id -u)" "$PLIST_DST" 2>/dev/null; then
  echo "launchd agent loaded."
else
  # Older macOS fallback.
  launchctl load -w "$PLIST_DST" && echo "launchd agent loaded (legacy load)."
fi

echo
echo "== probe =="
python3 probe.py --out probe-report.json || true
echo "Probe report: $REPO_DIR/probe-report.json"
echo "If the export below is wrong or empty, share that file — it has schema only, no event titles."

echo
echo "== first sync =="
./run.sh || true
tail -n 5 "$HOME/Library/Logs/cal-drop.log" 2>/dev/null || true
echo
echo "Done. Check the repo's data branch for calendar.json, then look at"
echo "~/Library/Logs/cal-drop.log and .local/last-run.log if anything failed."
