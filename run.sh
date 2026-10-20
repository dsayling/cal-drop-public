#!/bin/bash
# cal-drop launcher for launchd. launchd gives a minimal environment, so
# set PATH explicitly (git, ssh, python3 must all resolve), then run sync.
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
mkdir -p "$HOME/Library/Logs" "$REPO_DIR/.local"
cd "$REPO_DIR" || exit 1
PYTHON="$(command -v python3 || echo /usr/bin/python3)"
"$PYTHON" "$REPO_DIR/sync.py" >> "$HOME/Library/Logs/cal-drop.log" 2>&1
