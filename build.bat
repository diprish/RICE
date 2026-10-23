@echo off
rem Build a double-click offline bundle for Windows.
rem Output: dist\RICE-Tracker\ (zip it and share). Needs Python 3 on PATH.
setlocal
cd /d "%~dp0"

if not exist .venv\Scripts\python.exe (
    python -m venv .venv || goto :fail
)
set PY=.venv\Scripts\python.exe

%PY% -m pip install --quiet -r requirements.txt pyinstaller || goto :fail
%PY% -m PyInstaller --noconfirm --clean --onedir --name RICE-Tracker ^
  --add-data "templates;templates" --add-data "static;static" app.py || goto :fail

rem Ship the current workbook as a first-run seed; live data lives in %%APPDATA%%\RICE-Tracker.
if not exist dist\RICE-Tracker\seed mkdir dist\RICE-Tracker\seed
if exist rice_tracker_data.xlsx copy /y rice_tracker_data.xlsx dist\RICE-Tracker\seed\ >nul

echo Built dist\RICE-Tracker - zip that folder to share.
exit /b 0

:fail
echo Build failed.
exit /b 1
