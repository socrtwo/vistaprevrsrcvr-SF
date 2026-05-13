#!/usr/bin/env bash
# ChromeOS launcher for VistaPrevRsrcvr.
# Requires the Linux (Crostini) container: Settings > Advanced > Developers >
# Linux development environment. Once enabled, drop this bundle in your Linux
# files area and run this script from the Terminal.
set -euo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
EXE="$(find "$DIR" -maxdepth 1 -iname '*.exe' | head -n1)"

if [ -z "${EXE:-}" ]; then
  echo "Error: no .exe found in $DIR" >&2
  exit 1
fi

if ! command -v mono >/dev/null 2>&1; then
  cat >&2 <<'EOF'
Mono runtime is required inside the ChromeOS Linux container.
Install with:
  sudo apt-get update && sudo apt-get install -y mono-complete
EOF
  exit 127
fi

exec mono "$EXE" "$@"
