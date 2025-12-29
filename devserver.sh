#!/bin/bash
set -e

# Hapus venv jika rusak (opsional, tapi aman)
# rm -rf .venv

if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
fi

# Aktivasi venv
source .venv/bin/activate

# Upgrade pip
pip install --upgrade pip

# Install dependencies
echo "Installing dependencies..."
pip install -r requirements.txt

# Run the server
echo "Starting server..."
uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000} --reload
