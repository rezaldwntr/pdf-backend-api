# Gunakan image Python yang ringan
FROM python:3.9-slim

# --- BAGIAN YANG DIPERBAIKI ---
# Di Linux versi baru, 'libgl1-mesa-glx' diganti menjadi 'libgl1'
RUN apt-get update && apt-get install -y \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*
# -----------------------------

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Expose port
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
