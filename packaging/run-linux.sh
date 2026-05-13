#!/usr/bin/env bash
# Linux launcher for VistaPrevRsrcvr via Mono.
# VSS (Volume Shadow Copy Service) calls require a Windows host; on Linux the
# UI runs for browsing shadow-copy snapshots from mounted Windows volumes.
set -euo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
EXE="$(find "$DIR" -maxdepth 1 -iname '*.exe' | head -n1)"

if [ -z "${EXE:-}" ]; then
  echo "Error: no .exe found in $DIR" >&2
  exit 1
fi

if ! command -v mono >/dev/null 2>&1; then
  cat >&2 <<'EOF'
Mono runtime is required. Install with one of:
  Debian/Ubuntu: sudo apt-get install mono-complete
  Fedora:        sudo dnf install mono-complete
  Arch:          sudo pacman -S mono
EOF
  exit 127
fi

exec mono "$EXE" "$@"
