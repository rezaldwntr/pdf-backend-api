# Gunakan image Python yang ringan
FROM python:3.9-slim

# --- UPDATE DEPENDENCY ---
# Kita tambahkan 'poppler-utils' di sini agar error pdf2image hilang
RUN apt-get update && apt-get install -y \
    libgl1 \
    libglib2.0-0 \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*
# -------------------------

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Expose port
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
