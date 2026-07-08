#!/usr/bin/env bash
# Set up the Python virtual environment (if needed) and run the RICE tracker app.
set -euo pipefail
cd "$(dirname "$0")"

VENV_DIR=".venv"
VENV_PY="$VENV_DIR/bin/python"

# Recreate the venv if it's missing or unusable (venvs break when the
# project folder is moved or renamed).
if ! "$VENV_PY" -m pip --version >/dev/null 2>&1; then
    echo "Creating virtual environment in $VENV_DIR ..."
    rm -rf "$VENV_DIR"
    python3 -m venv "$VENV_DIR"
fi

echo "Installing dependencies ..."
"$VENV_PY" -m pip install --quiet --upgrade pip
"$VENV_PY" -m pip install --quiet -r requirements.txt

echo "Starting RICE Delivery Tracker at http://127.0.0.1:5000 ..."
exec "$VENV_PY" app.py
