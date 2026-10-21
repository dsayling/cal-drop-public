#!/bin/bash
# cal-drop launcher for launchd. launchd gives a minimal environment, so
# set PATH explicitly (git, ssh, python3 must all resolve), then run sync.
export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
mkdir -p "$HOME/Library/Logs" "$REPO_DIR/.local"
cd "$REPO_DIR" || exit 1
# Prefer cal-drop's dedicated interpreter (install.sh puts a uv-managed
# Python under .local/uv-python, used by nothing else). It exists so
# macOS Full Disk Access can be granted to that one binary instead of
# the system python3: grants follow the resolved executable path, and
# a venv's python is only a symlink back to the main interpreter.
PYTHON=""
for cand in "$REPO_DIR"/.local/uv-python/cpython-*/bin/python3; do
  if [ -x "$cand" ]; then PYTHON="$cand"; break; fi
done
if [ -z "$PYTHON" ]; then
  PYTHON="$(command -v python3 || echo /usr/bin/python3)"
fi
"$PYTHON" "$REPO_DIR/sync.py" >> "$HOME/Library/Logs/cal-drop.log" 2>&1
