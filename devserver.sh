#!/bin/bash
if [ ! -d ".venv" ]; then
    python3 -m venv .venv
fi
source .venv/bin/activate

# Install dependencies if not installed
if ! pip show fastapi > /dev/null; then
    pip install -r requirements.txt
fi

# Run the server
uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000} --reload
