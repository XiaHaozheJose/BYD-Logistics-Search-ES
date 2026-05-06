#!/bin/bash
set -e

cd "$(dirname "$0")"

echo "=== BYD Logistics Search — Local Mode ==="
echo

# Check Python
PYTHON=""
for p in python3.12 python3.11 python3.10 python3; do
  if command -v "$p" &>/dev/null; then
    PYTHON="$p"
    break
  fi
done

if [ -z "$PYTHON" ]; then
  echo "ERROR: Python 3.10+ not found. Please install Python."
  exit 1
fi

echo "Using: $PYTHON ($($PYTHON --version))"

# Install dependencies if needed
if ! $PYTHON -c "import flask, openpyxl" 2>/dev/null; then
  echo "Installing dependencies..."
  $PYTHON -m pip install --break-system-packages flask openpyxl 2>/dev/null || \
  $PYTHON -m pip install flask openpyxl
fi

echo
echo "Starting local server..."
echo "Open http://localhost:5000 in your browser"
echo "No login required — all data stays on your machine"
echo "Press Ctrl+C to stop"
echo

LOCAL_MODE=1 $PYTHON app.py
