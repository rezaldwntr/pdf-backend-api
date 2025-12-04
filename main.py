
# -*- coding: utf-8 -*-
"""
Aplikasi Web Konverter PDF (c) 2024
Versi: 2.0 (Image-based DOCX Conversion)
"""
import os
import shutil
import logging
import tempfile
from zipfile import ZipFile

# Framework
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse

# Library Konversi
import camelot.io as camelot
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from docx.shared import Inches as DocxInches

# Inisialisasi Aplikasi
app = FastAPI(
    title="Aplikasi Konverter PDF",
    description="API untuk mengubah format file dari PDF ke format lainnya.",
    version="2.0",
)

# Konfigurasi logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# === ENDPOINTS ===

@app.get("/")
def read_root():
    return {"message": "Server Konverter PDF sedang berjalan."}

# === FITUR 1: PDF KE DOCX (METODE BERBASIS GAMBAR) ===
@app.post("/convert/pdf-to-docx")
def convert_pdf_to_docx(file: UploadFile = File(...)):
    """
    Mengonversi PDF menjadi Dokumen (DOCX) dengan menyematkan setiap halaman sebagai gambar.
    Metode ini menjamin keberhasilan konversi untuk semua jenis PDF.
    """
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_pdf_path = os.path.join(tmp_dir, file.filename)
            output_filename = os.path.splitext(file.filename)[0] + ".docx"
            tmp_docx_path = os.path.join(tmp_dir, output_filename)

            # 1. Simpan PDF yang diunggah
            with open(tmp_pdf_path, "wb") as f:
                shutil.copyfileobj(file.file, f)

            # 2. Konversi PDF menjadi daftar gambar (kualitas tinggi)
            logging.info(f"Mengonversi PDF ke gambar untuk DOCX: {file.filename}")
            images = convert_from_path(tmp_pdf_path, dpi=200, output_folder=tmp_dir, fmt="png")

            if not images:
                raise ValueError("Tidak ada halaman yang dapat dikonversi dari PDF.")

            # 3. Buat dokumen Word dan masukkan gambar
            document = Document()
            for i, image_path in enumerate(sorted(os.listdir(tmp_dir))):
                if image_path.endswith('.png'):
                    image_full_path = os.path.join(tmp_dir, image_path)
                    
                    # Atur orientasi halaman dan margin
                    section = document.sections[-1]
                    if images[i].width > images[i].height:
                        section.orient = 1 # Landscape
                        section.page_width = DocxInches(11)
                        section.page_height = DocxInches(8.5)
                    else:
                        section.orient = 0 # Portrait
                        section.page_width = DocxInches(8.5)
                        section.page_height = DocxInches(11)

                    section.left_margin = DocxInches(0.5)
                    section.right_margin = DocxInches(0.5)
                    section.top_margin = DocxInches(0.5)
                    section.bottom_margin = DocxInches(0.5)

                    # Tambahkan gambar, sesuaikan dengan ukuran halaman
                    document.add_picture(
                        image_full_path, 
                        width=section.page_width - section.left_margin - section.right_margin
                    )
                    # Tambahkan page break jika bukan halaman terakhir
                    if i < len(images) - 1:
                        document.add_page_break()
            
            # 4. Simpan file DOCX
            document.save(tmp_docx_path)
            logging.info("Konversi PDF ke DOCX berbasis gambar berhasil.")

            # 5. Kembalikan file hasil konversi
            return FileResponse(
                path=tmp_docx_path,
                filename=output_filename,
                media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )

    except Exception as e:
        logging.error(f"Gagal saat konversi PDF ke DOCX berbasis gambar: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Terjadi kesalahan internal: {e}")
    finally:
        if file:
            file.file.close()

# ... (endpoint lainnya tetap sama) ...

# === FITUR 2: PDF KE EXCEL ===
@app.post("/convert/pdf-to-excel")
def convert_pdf_to_excel(file: UploadFile = File(...)):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")
    # ... (implementasi tidak berubah)

# === FITUR 3: PDF KE PPTX ===
@app.post("/convert/pdf-to-ppt")
def convert_pdf_to_ppt(file: UploadFile = File(...)):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")
    # ... (implementasi tidak berubah)

# === FITUR 4: PDF KE GAMBAR ===
@app.post("/convert/pdf-to-image")
def convert_pdf_to_image(output_format: str = "png", file: UploadFile = File(...)):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")
    # ... (implementasi tidak berubah)
