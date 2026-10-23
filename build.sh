#!/usr/bin/env bash
# Build a double-click offline bundle for the CURRENT OS (PyInstaller can't cross-compile).
# Output: dist/RICE-Tracker/ (zip it and share). On Windows use build.bat.
set -euo pipefail
cd "$(dirname "$0")"
[ -x .venv/bin/python ] || python3 -m venv .venv
PY=.venv/bin/python
$PY -m pip install --quiet -r requirements.txt pyinstaller
$PY -m PyInstaller --noconfirm --clean --onedir --name RICE-Tracker \
  --add-data "templates:templates" --add-data "static:static" app.py
# Ship the current workbook as a first-run seed. Live data (saved filters, uploads) lives in the
# per-user data folder, so rebuilding or replacing this folder never touches it.
mkdir -p dist/RICE-Tracker/seed
[ -f rice_tracker_data.xlsx ] && cp rice_tracker_data.xlsx dist/RICE-Tracker/seed/
echo "Built dist/RICE-Tracker — zip that folder to share."
