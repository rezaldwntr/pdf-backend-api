# 1. PILIH BASE IMAGE
# Kita menggunakan Python versi 3.9 versi "slim" (ringan)
# agar proses download dan deploy lebih cepat dan hemat memori.
FROM python:3.9-slim

# 2. INSTALL DEPENDENCY SISTEM (LINUX)
# Bagian ini sangat KRUSIAL. Kita menginstall library tambahan yang dibutuhkan oleh Python.
# - libgl1 & libglib2.0-0: Dibutuhkan oleh OpenCV (untuk pemrosesan gambar).
# - poppler-utils: WAJIB ADA untuk library 'pdf2image' agar bisa membaca info halaman PDF.
# - rm -rf ...: Membersihkan cache apt agar ukuran image tetap kecil.
RUN apt-get update && apt-get install -y \
    libgl1 \
    libglib2.0-0 \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# 3. SET FOLDER KERJA
# Membuat folder bernama '/app' di dalam container dan menjadikannya folder aktif.
WORKDIR /app

# 4. INSTALL LIBRARY PYTHON
# Pertama, kita copy dulu file daftar library (requirements.txt) ke dalam container.
COPY requirements.txt .
# Lalu kita suruh pip untuk menginstall semua library yang terdaftar di situ.
# --no-cache-dir digunakan agar tidak menyimpan file mentahan (hemat space).
RUN pip install --no-cache-dir -r requirements.txt

# 5. COPY SISA KODE
# Mengcopy semua file kodingan kita (main.py, dll) dari laptop ke dalam folder /app di container.
COPY . .

# 6. JALANKAN SERVER
# Perintah ini yang akan dijalankan saat server menyala.
# uvicorn: Nama servernya.
# main:app : Menjalankan objek 'app' yang ada di file 'main.py'.
# --host 0.0.0.0: Agar server bisa diakses dari luar container.
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
