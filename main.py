# -*- coding: utf-8 -*-
"""
Aplikasi Web Konverter PDF (c) 2024
Versi: 2.2 (Editable PPT & Better Excel Logic)
"""
import os
import shutil
import logging
import tempfile
from zipfile import ZipFile

# Framework
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

# Library Konversi
from pdf2docx import Converter
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches
from docx import Document
from docx.shared import Inches as DocxInches
import pdfplumber
import pandas as pd

# Konfigurasi
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10 MB

# Inisialisasi Aplikasi
app = FastAPI(
    title="Aplikasi Konverter PDF",
    description="API untuk mengubah format file dari PDF ke format lainnya.",
    version="2.2",
)

# Helper untuk hapus folder (jika belum ada)
def cleanup_folder(path: str):
    try:
        if os.path.exists(path):
            shutil.rmtree(path)
            logging.info(f"Deleted temp folder: {path}")
    except Exception as e:
        logging.error(f"Error cleaning up: {e}")

# Helper validasi file
def validate_file(file: UploadFile):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")
    
    file.file.seek(0, 2)
    file_size = file.file.tell()
    file.file.seek(0)
    
    if file_size > MAX_FILE_SIZE:
        raise HTTPException(status_code=400, detail=f"Ukuran file terlalu besar (Maks {MAX_FILE_SIZE/1024/1024}MB)")

@app.get("/")
def read_root():
    return {"message": "Server PDF Backend (Editable Version) is Running!"}

# === FITUR 1: PDF KE DOCX ===
@app.post("/convert/pdf-to-docx")
def convert_pdf_to_docx(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    validate_file(file)

    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    docx_filename = os.path.splitext(file.filename)[0] + ".docx"
    tmp_docx_path = os.path.join(tmp_dir, docx_filename)

    try:
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Konversi PDF ke DOCX
        # Menggunakan multiprocess=False untuk stabilitas CPU rendah
        cv = Converter(tmp_pdf_path)
        cv.convert(tmp_docx_path, start=0, end=None, multiprocess=False)
        cv.close()

        background_tasks.add_task(cleanup_folder, tmp_dir)

        return FileResponse(
            path=tmp_docx_path,
            filename=docx_filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error(f"ERROR PDF TO DOCX: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Gagal convert Word: {str(e)}")

# === FITUR 2: PDF KE EXCEL (DIPERBAIKI: Gabung Sheet) ===
@app.post("/convert/pdf-to-excel")
def convert_pdf_to_excel(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    validate_file(file)
    
    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    xlsx_filename = os.path.splitext(file.filename)[0] + ".xlsx"
    tmp_xlsx_path = os.path.join(tmp_dir, xlsx_filename)

    try:
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Logika ekstraksi tabel yang diperbarui (Menggabungkan semua tabel)
        with pdfplumber.open(tmp_pdf_path) as pdf:
            with pd.ExcelWriter(tmp_xlsx_path, engine='openpyxl') as writer:
                all_rows = []
                
                # Kumpulkan semua data dari semua halaman terlebih dahulu
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if table:
                            # Bersihkan data None menjadi string kosong
                            clean_table = [[cell if cell is not None else "" for cell in row] for row in table]
                            all_rows.extend(clean_table)
                            # Tambahkan baris kosong sebagai pemisah antar tabel
                            all_rows.append([]) 

                if all_rows:
                    df = pd.DataFrame(all_rows)
                    # Simpan ke satu sheet utama agar mudah diedit
                    df.to_excel(writer, sheet_name="Hasil Konversi", index=False, header=False)
                else:
                    # Jika tidak ada tabel ditemukan
                    pd.DataFrame(["Tidak ada tabel ditemukan dalam PDF"]).to_excel(writer, sheet_name="Info", index=False, header=False)

        background_tasks.add_task(cleanup_folder, tmp_dir)

        return FileResponse(
            path=tmp_xlsx_path,
            filename=xlsx_filename,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error(f"ERROR PDF TO EXCEL: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Gagal convert Excel: {str(e)}")

# === FITUR 3: PDF KE PPTX (DIPERBAIKI: Text Box Editable) ===
@app.post("/convert/pdf-to-ppt")
def convert_pdf_to_ppt(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    validate_file(file)
    
    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    ppt_filename = os.path.splitext(file.filename)[0] + ".pptx"
    tmp_ppt_path = os.path.join(tmp_dir, ppt_filename)

    try:
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Konversi PDF ke PPTX (Metode Text Extraction agar Editable)
        prs = Presentation()
        doc = fitz.open(tmp_pdf_path)

        for page_num, page in enumerate(doc):
            # 1. Buat slide kosong (Layout 6 = Blank)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            
            # 2. Ambil blok teks dari halaman PDF
            # "dict" memberikan koordinat dan isi teks
            blocks = page.get_text("dict")["blocks"]
            
            for b in blocks:
                if b['type'] == 0: # Tipe 0 adalah Teks
                    for line in b["lines"]:
                        for span in line["spans"]:
                            text = span["text"]
                            # Skip teks kosong
                            if not text.strip():
                                continue
                                
                            # Ambil koordinat (x, y) dari PDF
                            x, y = span["bbox"][:2]
                            
                            # Konversi koordinat PDF (points) ke PPTX (Inches)
                            # 1 Inch = 72 Points (standar PDF)
                            left = Inches(x / 72)
                            top = Inches(y / 72)
                            
                            # Ukuran Textbox (lebar default 5 inci, tinggi menyesuaikan)
                            width = Inches(5) 
                            height = Inches(0.5)

                            # Tambahkan Textbox ke slide
                            txBox = slide.shapes.add_textbox(left, top, width, height)
                            tf = txBox.text_frame
                            tf.text = text
                            
                            # Opsional: Jika ingin menambahkan gambar juga, logika type==1 bisa ditambahkan disini
                            # Namun fokus saat ini adalah teks yang bisa diedit.

        doc.close()
        prs.save(tmp_ppt_path)

        background_tasks.add_task(cleanup_folder, tmp_dir)

        return FileResponse(
            path=tmp_ppt_path,
            filename=ppt_filename,
            media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error(f"ERROR PDF TO PPT: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Gagal convert PPT: {str(e)}")

# === FITUR 4: PDF KE GAMBAR (ZIP) ===
@app.post("/convert/pdf-to-image")
def convert_pdf_to_image(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    output_format: str = "png"
):
    validate_file(file)
    
    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    zip_filename = os.path.splitext(file.filename)[0] + "_images.zip"
    tmp_zip_path = os.path.join(tmp_dir, zip_filename)

    try:
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = fitz.open(tmp_pdf_path)
        
        with ZipFile(tmp_zip_path, 'w') as zipf:
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(dpi=200) # DPI lebih tinggi untuk kualitas gambar
                
                img_name = f"page_{page_num + 1}.{output_format}"
                img_path = os.path.join(tmp_dir, img_name)
                
                # Support format output sederhana
                if output_format.lower() in ['jpg', 'jpeg']:
                    pix.save(img_path, output="jpg")
                else:
                    pix.save(img_path)
                
                zipf.write(img_path, img_name)

        doc.close()

        background_tasks.add_task(cleanup_folder, tmp_dir)

        return FileResponse(
            path=tmp_zip_path,
            filename=zip_filename,
            media_type='application/zip'
        )

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error(f"ERROR PDF TO IMAGE: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Gagal convert Image: {str(e)}")
