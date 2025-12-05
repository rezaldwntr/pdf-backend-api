# -*- coding: utf-8 -*-
# --- BAGIAN 1: IMPORT LIBRARY ---
import os
import shutil
import logging
import tempfile
import traceback
import zipfile
import io

# Framework FastAPI: Untuk membuat API Server
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

# Library Konversi PDF:
from pdf2docx import Converter  # Untuk ubah PDF ke Word
import fitz  # PyMuPDF (Sangat cepat untuk membaca PDF & ekstrak gambar)
from pptx import Presentation # Untuk membuat file PowerPoint
from pptx.util import Inches
import pdfplumber # Untuk ekstrak tabel dari PDF
import pandas as pd # Untuk mengolah data tabel jadi Excel

# --- BAGIAN 2: KONFIGURASI APLIKASI ---

# Inisialisasi aplikasi FastAPI
app = FastAPI(
    title="Aplikasi Konverter PDF",
    description="Backend API untuk mengubah PDF ke Word, Excel, PPT, dan Gambar.",
    version="2.1",
)

# Konfigurasi CORS (Cross-Origin Resource Sharing)
# Ini PENTING agar frontend (Vercel) diizinkan mengakses backend ini.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # Mengizinkan semua domain mengakses API ini
    allow_methods=["*"], # Mengizinkan semua metode (GET, POST, dll)
    allow_headers=["*"], # Mengizinkan semua header request
)

# Konfigurasi Logging: Agar kita bisa melihat pesan error/info di dashboard Render
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Fungsi Helper: Membersihkan Folder Sementara
# Fungsi ini akan dipanggil otomatis setelah file selesai dikirim ke user.
def cleanup_folder(path: str):
    try:
        if os.path.exists(path):
            shutil.rmtree(path) # Hapus folder beserta isinya
            logging.info(f"Berhasil menghapus folder temp: {path}")
    except Exception as e:
        logging.error(f"Gagal menghapus folder: {e}")

# Endpoint Root: Untuk mengecek apakah server hidup
@app.get("/")
def read_root():
    return {"message": "Server Konverter PDF sedang berjalan dengan normal."}


# --- BAGIAN 3: ENDPOINT KONVERSI ---

# === FITUR 1: PDF KE WORD (DOCX) ===
@app.post("/convert/pdf-to-docx")
async def convert_pdf_to_docx(
    background_tasks: BackgroundTasks, # Tugas latar belakang untuk bersih-bersih
    file: UploadFile = File(...) # File yang diupload user
):
    # Validasi: Pastikan yang diupload adalah PDF
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    # Persiapan Folder Sementara (Temp)
    # mkdtemp() membuat folder dengan nama acak agar tidak bentrok antar user
    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    docx_filename = os.path.splitext(file.filename)[0] + ".docx"
    tmp_docx_path = os.path.join(tmp_dir, docx_filename)

    try:
        # 1. Simpan file PDF dari user ke folder temp
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # 2. Proses Konversi (PDF -> DOCX)
        # Menggunakan library pdf2docx
        cv = Converter(tmp_pdf_path)
        # multiprocess=False digunakan agar tidak memberatkan CPU server gratisan
        cv.convert(tmp_docx_path, start=0, end=None, multiprocess=False, cpu_count=1)
        cv.close()

        # 3. Jadwalkan penghapusan folder temp SETELAH file terkirim
        background_tasks.add_task(cleanup_folder, tmp_dir)

        # 4. Kirim file hasil (DOCX) ke user
        return FileResponse(
            path=tmp_docx_path,
            filename=docx_filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        # Jika terjadi error, hapus folder temp sekarang juga
        cleanup_folder(tmp_dir)
        logging.error("!!! ERROR PDF TO DOCX !!!")
        traceback.print_exc() # Cetak detail error di log server
        raise HTTPException(status_code=500, detail=f"Gagal convert Word: {str(e)}")


# === FITUR 2: PDF KE EXCEL (XLSX) ===
@app.post("/convert/pdf-to-excel")
async def convert_pdf_to_excel(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    # Setup path sementara
    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    xlsx_filename = os.path.splitext(file.filename)[0] + ".xlsx"
    tmp_xlsx_path = os.path.join(tmp_dir, xlsx_filename)

    try:
        # Simpan PDF
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Logic Konversi: Cari tabel di PDF
        with pdfplumber.open(tmp_pdf_path) as pdf:
            # Siapkan writer Excel
            with pd.ExcelWriter(tmp_xlsx_path, engine='openpyxl') as writer:
                for i, page in enumerate(pdf.pages):
                    # Ekstrak tabel
                    tables = page.extract_tables()
                    if tables:
                        # Jika ada tabel, tulis ke sheet Excel
                        for table in tables:
                            df = pd.DataFrame(table)
                            df.to_excel(writer, sheet_name=f"Page {i+1}", index=False, header=False)
                    else:
                        # Jika tidak ada tabel, tulis info kosong
                        pd.DataFrame(["Tidak ditemukan tabel"]).to_excel(writer, sheet_name=f"Page {i+1}")

        # Jadwalkan bersih-bersih
        background_tasks.add_task(cleanup_folder, tmp_dir)

        # Kirim File
        return FileResponse(
            path=tmp_xlsx_path,
            filename=xlsx_filename,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error("!!! ERROR PDF TO EXCEL !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Excel: {str(e)}")


# === FITUR 3: PDF KE POWERPOINT (PPTX) ===
@app.post("/convert/pdf-to-ppt")
async def convert_pdf_to_ppt(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    # Setup path sementara
    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    ppt_filename = os.path.splitext(file.filename)[0] + ".pptx"
    tmp_ppt_path = os.path.join(tmp_dir, ppt_filename)
    
    try:
        # Simpan PDF
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Logic: PDF -> Gambar -> Slide PPT
        prs = Presentation()
        # Set ukuran slide ke Widescreen (16:9)
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        # Buka PDF dengan PyMuPDF
        doc = fitz.open(tmp_pdf_path)

        for i, page in enumerate(doc):
            # Buat slide kosong
            blank_slide_layout = prs.slide_layouts[6] 
            slide = prs.slides.add_slide(blank_slide_layout)
            
            # Ubah halaman PDF jadi Gambar (PNG)
            pix = page.get_pixmap(dpi=150)
            img_filename = os.path.join(tmp_dir, f"slide_{i}.png")
            pix.save(img_filename)
            
            # Tempel gambar ke slide (Full Screen)
            slide.shapes.add_picture(img_filename, 0, 0, width=prs.slide_width, height=prs.slide_height)

        # Simpan file PPT
        prs.save(tmp_ppt_path)
        doc.close()

        # Jadwalkan bersih-bersih
        background_tasks.add_task(cleanup_folder, tmp_dir)

        return FileResponse(
            path=tmp_ppt_path,
            filename=ppt_filename,
            media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error("!!! ERROR PDF TO PPT !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert PPT: {str(e)}")


# === FITUR 4: PDF KE GAMBAR (ZIP) ===
@app.post("/convert/pdf-to-image")
async def convert_pdf_to_image(
    background_tasks: BackgroundTasks,
    output_format: str = "png", # User bisa memilih jpg/png
    file: UploadFile = File(...)
):
    # Normalisasi format ke huruf kecil
    fmt = output_format.lower()
    if fmt not in ["jpg", "jpeg", "png"]:
        raise HTTPException(status_code=400, detail="Format tidak didukung. Pilih: jpg, jpeg, atau png.")

    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    # Setup path sementara dan folder khusus gambar
    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    zip_filename = os.path.splitext(file.filename)[0] + f"_images_{fmt}.zip"
    tmp_zip_path = os.path.join(tmp_dir, zip_filename)
    
    # Buat sub-folder untuk menampung gambar
    images_folder = os.path.join(tmp_dir, "processed_images")
    os.makedirs(images_folder, exist_ok=True)

    try:
        # Simpan PDF
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = fitz.open(tmp_pdf_path)
        
        # Tentukan apakah pakai background transparan (Alpha)
        # JPG tidak support transparan, jadi False. PNG support, jadi True.
        use_alpha = False if fmt in ["jpg", "jpeg"] else True

        # Loop setiap halaman
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=150, alpha=use_alpha)
            # Nama file gambar: page_001.png, page_002.png...
            img_path = os.path.join(images_folder, f"page_{str(i+1).zfill(3)}.{fmt}")
            pix.save(img_path)

        doc.close()

        # Kompres semua gambar ke dalam file ZIP
        with zipfile.ZipFile(tmp_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(images_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    # arcname agar struktur folder di dalam zip rapi
                    zipf.write(file_path, arcname=file)

        # Jadwalkan bersih-bersih
        background_tasks.add_task(cleanup_folder, tmp_dir)

        # Kirim file ZIP
        return FileResponse(
            path=tmp_zip_path,
            filename=zip_filename,
            media_type='application/zip'
        )

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error("!!! ERROR PDF TO IMAGE !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Image: {str(e)}")
