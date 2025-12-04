
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
from fastapi import BackgroundTasks
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

# Helper untuk hapus folder (jika belum ada)
def cleanup_folder(path: str):
    try:
        if os.path.exists(path):
            shutil.rmtree(path)
    except: pass

# Konfigurasi logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# === ENDPOINTS ===

@app.get("/")
def read_root():
    return {"message": "Server Konverter PDF sedang berjalan."}

# === UPDATE FITUR PDF KE DOCX ===
@app.post("/convert/pdf-to-docx")
async def convert_pdf_to_docx(
    background_tasks: BackgroundTasks,  # <--- Tambah ini
    file: UploadFile = File(...)
):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    # Gunakan folder temp unik agar aman
    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    
    # Nama output docx
    docx_filename = os.path.splitext(file.filename)[0] + ".docx"
    tmp_docx_path = os.path.join(tmp_dir, docx_filename)

    try:
        # Simpan PDF
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Proses Konversi
        cv = Converter(tmp_pdf_path)
        cv.convert(tmp_docx_path, start=0, end=None, multiprocess=False, cpu_count=1)
        cv.close()

        # Jadwalkan Hapus Folder SETELAH file terkirim
        background_tasks.add_task(cleanup_folder, tmp_dir)

        # Kirim File
        return FileResponse(
            path=tmp_docx_path,
            filename=docx_filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        cleanup_folder(tmp_dir) # Hapus jika error
        print("!!! ERROR PDF TO DOCX !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Word: {str(e)}")

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
