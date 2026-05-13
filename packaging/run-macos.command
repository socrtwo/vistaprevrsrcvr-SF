#!/usr/bin/env bash
# macOS launcher for VistaPrevRsrcvr via Mono.
# Double-click in Finder to run. VSS operations require a Windows volume; on
# macOS this UI is useful for browsing shadow copies from mounted NTFS drives.
set -euo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
EXE="$(find "$DIR" -maxdepth 1 -iname '*.exe' | head -n1)"

if [ -z "${EXE:-}" ]; then
  osascript -e 'display alert "VistaPrevRsrcvr" message "No .exe found next to this launcher."'
  exit 1
fi

if ! command -v mono >/dev/null 2>&1; then
  osascript -e 'display alert "Mono required" message "Install Mono from https://www.mono-project.com/download/stable/ (or: brew install mono)."'
  exit 127
fi

exec mono "$EXE" "$@"
