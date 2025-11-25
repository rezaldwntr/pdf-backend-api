# Gunakan Python versi ringan
FROM python:3.9-slim

# Install library sistem untuk grafis (Wajib untuk library PDF/Image processing)
RUN apt-get update && apt-get install -y \
    libgl1-mesa-glx \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# Set folder kerja
WORKDIR /app

# Copy dan install requirements
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy sisa kode
COPY . .

# Jalankan server pada port 8000
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
