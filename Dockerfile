# 1. PILIH BASE IMAGE
FROM python:3.9-slim

# 2. INSTALL DEPENDENCY SISTEM (LINUX)
# Menggabungkan update, install, dan clean dalam satu layer untuk mengurangi ukuran image
RUN apt-get update && apt-get install -y \
    libgl1 \
    libglib2.0-0 \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# 3. SET FOLDER KERJA
WORKDIR /app

# 4. INSTALL LIBRARY PYTHON
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 5. COPY SISA KODE
# (Docker akan mengabaikan file yang ada di .dockerignore)
COPY . .

# 6. JALANKAN SERVER
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
