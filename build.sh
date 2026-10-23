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
# Seed the data folder with the current workbook so recipients open straight to the dashboard.
mkdir -p dist/RICE-Tracker/data
[ -f rice_tracker_data.xlsx ] && cp rice_tracker_data.xlsx dist/RICE-Tracker/data/
echo "Built dist/RICE-Tracker — zip that folder to share."
