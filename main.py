
# -*- coding: utf-8 -*-
"""
Aplikasi Web Konverter PDF (c) 2024
Versi: 1.2 (PikePDF Fix)
"""
import os
import shutil
import traceback
import logging
import tempfile
from typing import List
from zipfile import ZipFile

# Framework
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, JSONResponse

# Library Konversi
from pdf2docx import Converter as cv_docx
import camelot.io as camelot
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches
import pikepdf  # <-- Dependensi baru untuk perbaikan PDF

# Inisialisasi Aplikasi
app = FastAPI(
    title="Aplikasi Konverter PDF",
    description="API untuk mengubah format file dari PDF ke format lainnya.",
    version="1.2",
)

# Konfigurasi logging terpusat
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# === ENDPOINTS ===

@app.get("/")
def read_root():
    return {"message": "Server Konverter PDF sedang berjalan."}

# === FITUR 1: PDF KE DOCX (DENGAN PERBAIKAN) ===
@app.post("/convert/pdf-to-docx")
def convert_pdf_to_docx(file: UploadFile = File(...)):
    """
    Mengonversi file PDF menjadi Dokumen (DOCX).
    Menggunakan `pdf2docx` dengan pra-proses perbaikan oleh `pikepdf`.
    """
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            # Path sementara di dalam direktori temporer
            original_pdf_path = os.path.join(tmp_dir, file.filename)
            repaired_pdf_path = os.path.join(tmp_dir, "repaired.pdf") # Path untuk PDF yang sudah bersih
            output_filename = os.path.splitext(file.filename)[0] + ".docx"
            tmp_docx_path = os.path.join(tmp_dir, output_filename)

            # 1. Simpan file PDF yang diupload
            with open(original_pdf_path, "wb") as f:
                shutil.copyfileobj(file.file, f)

            # 2. [PERBAIKAN] Bersihkan/Perbaiki PDF menggunakan pikepdf
            logging.info(f"Memperbaiki PDF: {file.filename}")
            with pikepdf.open(original_pdf_path) as pdf:
                pdf.save(repaired_pdf_path)
            logging.info("Perbaikan PDF selesai.")

            # 3. Proses Konversi menggunakan PDF yang sudah diperbaiki
            cv = cv_docx(repaired_pdf_path) # <-- Gunakan PDF yang bersih
            cv.convert(tmp_docx_path, start=0, end=None)
            cv.close()

            # 4. Kembalikan file hasil konversi
            return FileResponse(
                path=tmp_docx_path,
                filename=output_filename,
                media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
    except Exception as e:
        logging.error(f"Gagal saat konversi PDF ke DOCX: {e}", exc_info=True)
        # Memberikan detail error spesifik jika dari pikepdf
        if isinstance(e, pikepdf.PdfError):
            raise HTTPException(status_code=422, detail=f"Gagal memproses PDF. File mungkin rusak atau tidak valid: {e}")
        raise HTTPException(status_code=500, detail=f"Terjadi kesalahan internal: {e}")
    finally:
        # Pastikan file stream ditutup
        if file:
            file.file.close()


# === FITUR 2: PDF KE EXCEL ===
@app.post("/convert/pdf-to-excel")
def convert_pdf_to_excel(file: UploadFile = File(...)):
    """
    Mengonversi tabel dari PDF menjadi file Excel (XLSX).
    Menggunakan `camelot-py`.
    """
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_pdf_path = os.path.join(tmp_dir, file.filename)
            output_filename = os.path.splitext(file.filename)[0] + ".xlsx"
            tmp_excel_path = os.path.join(tmp_dir, output_filename)

            # 1. Simpan file PDF
            with open(tmp_pdf_path, "wb") as f:
                shutil.copyfileobj(file.file, f)

            # 2. Proses Ekstraksi Tabel
            tables = camelot.read_pdf(tmp_pdf_path, flavor='stream', pages='all')

            if tables.n == 0:
                raise HTTPException(status_code=404, detail="Tidak ada tabel yang terdeteksi di dalam PDF.")

            # 3. Simpan ke Excel
            tables.export(tmp_excel_path, f='excel', compress=False)
            
            actual_file_path = ""
            for item in os.listdir(tmp_dir):
                if item.endswith(".xlsx"):
                    actual_file_path = os.path.join(tmp_dir, item)
                    break
            
            if not actual_file_path:
                 raise FileNotFoundError("File Excel hasil konversi tidak ditemukan.")

            # 4. Kembalikan file
            return FileResponse(
                path=actual_file_path,
                filename=output_filename,
                media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    except Exception as e:
        logging.error(f"Gagal saat konversi PDF ke Excel: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Terjadi kesalahan internal: {e}")
    finally:
        if file:
            file.file.close()


# === FITUR 3: PDF KE PPTX ===
@app.post("/convert/pdf-to-ppt")
def convert_pdf_to_ppt(file: UploadFile = File(...)):
    """
    Mengonversi setiap halaman PDF menjadi slide di PowerPoint (PPTX).
    """
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_pdf_path = os.path.join(tmp_dir, file.filename)
            output_filename = os.path.splitext(file.filename)[0] + ".pptx"
            tmp_ppt_path = os.path.join(tmp_dir, output_filename)

            # 1. Simpan PDF
            with open(tmp_pdf_path, "wb") as f:
                shutil.copyfileobj(file.file, f)

            # 2. Konversi PDF ke list of images
            images = convert_from_path(tmp_pdf_path, output_folder=tmp_dir, fmt="png")

            # 3. Buat presentasi PowerPoint
            prs = Presentation()
            
            if not images:
                raise ValueError("Tidak ada halaman yang dapat dikonversi dari PDF.")

            first_page = images[0]
            prs.slide_width = Inches(first_page.width / 96)
            prs.slide_height = Inches(first_page.height / 96)

            for i, image_obj in enumerate(images):
                image_path = os.path.join(tmp_dir, f"page_{i}.png")
                image_obj.save(image_path, "PNG")

                blank_slide_layout = prs.slide_layouts[6]
                slide = prs.slides.add_slide(blank_slide_layout)
                slide.shapes.add_picture(image_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
            
            # 4. Simpan file PPTX
            prs.save(tmp_ppt_path)
            
            # 5. Kembalikan file
            return FileResponse(
                path=tmp_ppt_path,
                filename=output_filename,
                media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
    except Exception as e:
        logging.error(f"Gagal saat konversi PDF ke PPT: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Terjadi kesalahan internal: {e}")
    finally:
        if file:
            file.file.close()

# === FITUR 4: PDF KE GAMBAR (JPG, JPEG, PNG) ===
@app.post("/convert/pdf-to-image")
def convert_pdf_to_image(
    output_format: str = "png",
    file: UploadFile = File(...),
):
    """
    Mengonversi PDF menjadi file gambar. ZIP jika >1 halaman.
    """
    fmt = output_format.lower()
    if fmt not in ["jpg", "jpeg", "png"]:
        raise HTTPException(status_code=400, detail="Format tidak didukung. Pilih: jpg, jpeg, atau png.")
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_pdf_path = os.path.join(tmp_dir, file.filename)

            # 1. Simpan PDF
            with open(tmp_pdf_path, "wb") as f:
                shutil.copyfileobj(file.file, f)

            # 2. Konversi PDF ke gambar
            images = convert_from_path(tmp_pdf_path, output_folder=tmp_dir, fmt=fmt)
            
            image_paths = []
            for i, image_obj in enumerate(images):
                img_filename = f"page_{i+1}.{fmt}"
                img_path = os.path.join(tmp_dir, img_filename)
                image_obj.save(img_path, fmt.upper())
                image_paths.append(img_path)

            # 3. Tentukan output: file tunggal atau ZIP
            if len(image_paths) == 1:
                output_filename = os.path.splitext(file.filename)[0] + f".{fmt}"
                return FileResponse(path=image_paths[0], filename=output_filename, media_type=f"image/{fmt}")
            else:
                zip_filename = os.path.splitext(file.filename)[0] + ".zip"
                zip_path = os.path.join(tmp_dir, zip_filename)
                
                with ZipFile(zip_path, 'w') as zipf:
                    for img_path in image_paths:
                        zipf.write(img_path, os.path.basename(img_path))
                
                return FileResponse(path=zip_path, filename=zip_filename, media_type="application/zip")
    except Exception as e:
        logging.error(f"Gagal saat konversi PDF ke Gambar: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Terjadi kesalahan internal: {e}")
    finally:
        if file:
            file.file.close()

