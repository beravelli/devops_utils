#!/usr/bin/env bash
# Make openpyxl available for the planning scripts.
#
# Plain `pip install openpyxl` is tried first. On hosts where pip cannot resolve
# files.pythonhosted.org (seen on this machine) the wheels are fetched with curl
# from the URLs the PyPI JSON API reports, then installed from disk.
set -euo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/../../.." && pwd)"
VENV="$REPO/.venv"
PY="${PYTHON:-python3}"

if "$PY" -c 'import openpyxl' 2>/dev/null; then
  echo "openpyxl already available to $PY"; exit 0
fi

[ -d "$VENV" ] || "$PY" -m venv "$VENV"
if "$VENV/bin/python" -c 'import openpyxl' 2>/dev/null; then
  echo "openpyxl already available in $VENV"; exit 0
fi

if "$VENV/bin/pip" install -q openpyxl 2>/dev/null; then
  echo "installed openpyxl via pip into $VENV"
else
  echo "pip could not reach PyPI; falling back to curl" >&2
  WHEELS="$(mktemp -d)"
  for pkg in openpyxl et-xmlfile; do
    url="$(curl -fsS --max-time 20 "https://pypi.org/pypi/$pkg/json" | "$PY" -c '
import sys, json
d = json.load(sys.stdin); v = d["info"]["version"]
print(next(u["url"] for u in d["releases"][v] if u["url"].endswith(".whl")))')"
    echo "  fetching $(basename "$url")"
    curl -fsS --max-time 60 -o "$WHEELS/$(basename "$url")" "$url"
  done
  "$VENV/bin/pip" install -q --no-index --no-deps "$WHEELS"/*.whl
  rm -rf "$WHEELS"
  echo "installed openpyxl from downloaded wheels into $VENV"
fi

"$VENV/bin/python" -c 'import openpyxl; print("openpyxl", openpyxl.__version__)'
echo "use $VENV/bin/python to run the planning scripts"
