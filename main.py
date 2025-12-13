# -*- coding: utf-8 -*-
import os
import shutil
import logging
import tempfile
import traceback
import zipfile
import io

# Framework
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

# Library Konversi
from pdf2docx import Converter
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pptx.util import Pt

# Inisialisasi Aplikasi
app = FastAPI(
    title="Aplikasi Konverter PDF",
    version="3.1 (Optimized Version)",
)

# Setup CORS
# PENTING: Jika sudah ada domain frontend, ubah ["*"] jadi ["https://domain-kamu.com"]
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Konfigurasi
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10 MB

# Helper hapus folder
def cleanup_folder(path: str):
    try:
        if os.path.exists(path):
            shutil.rmtree(path)
            logging.info(f"Deleted temp folder: {path}")
    except Exception as e:
        logging.error(f"Error cleaning up: {e}")

# Helper validasi file
def validate_file(file: UploadFile):
    # Cek Ekstensi
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")
    
    # Cek Ukuran (Metode seek/tell aman untuk SpooledTemporaryFile)
    file.file.seek(0, 2)
    file_size = file.file.tell()
    file.file.seek(0)
    
    if file_size > MAX_FILE_SIZE:
        raise HTTPException(status_code=400, detail=f"Ukuran file terlalu besar (Maks {MAX_FILE_SIZE/1024/1024}MB)")

@app.get("/")
def read_root():
    return {"message": "Server Konverter PDF (Optimized) sedang berjalan."}

# === FITUR 1: PDF KE DOCX (EDITABLE) ===
# CATATAN: Menggunakan 'def' biasa (bukan async) karena ini proses berat (CPU bound).
# FastAPI akan otomatis menjalankannya di thread pool terpisah.
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

        cv = Converter(tmp_pdf_path)
        cv.convert(tmp_docx_path, start=0, end=None, multiprocess=False, cpu_count=1)
        cv.close()

        background_tasks.add_task(cleanup_folder, tmp_dir)

        return FileResponse(
            path=tmp_docx_path,
            filename=docx_filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error("!!! ERROR PDF TO DOCX !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Word: {str(e)}")


# === FITUR 2: PDF KE EXCEL (REVISI: DYNAMIC HEADER SEPARATION) ===
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

        # Siapkan Workbook Excel Manual
        wb = Workbook()
        ws = wb.active
        ws.title = "Hasil Convert"

        all_table_data = []
        header_text = []

        with pdfplumber.open(tmp_pdf_path) as pdf:
            # 1. LOGIKA BARU: DETEKSI BATAS ANTARA JUDUL DAN TABEL
            if len(pdf.pages) > 0:
                first_page = pdf.pages[0]
                
                # Cari tabel di halaman pertama untuk tahu posisinya
                found_tables = first_page.find_tables()
                
                if found_tables:
                    # Ambil koordinat bagian ATAS dari tabel pertama (bbox[1] adalah Y-top)
                    first_table_top = found_tables[0].bbox[1]
                    
                    # Ambil teks HANYA dari bagian atas halaman sampai batas atas tabel
                    # Kurangi sedikit (-5) agar garis tabel tidak ikut terbaca
                    safe_header_bottom = max(0, first_table_top - 5)
                    
                    header_crop = first_page.crop((0, 0, first_page.width, safe_header_bottom))
                    raw_header = header_crop.extract_text()
                else:
                    # Jika tidak ada tabel, coba ambil 20% teratas (fallback)
                    header_crop = first_page.crop((0, 0, first_page.width, first_page.height * 0.2))
                    raw_header = header_crop.extract_text()

                if raw_header:
                    header_text = raw_header.split('\n')

            # 2. AMBIL DATA TABEL DARI SEMUA HALAMAN
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    # Bersihkan data None menjadi string kosong
                    cleaned_table = [[cell if cell is not None else "" for cell in row] for row in table]
                    all_table_data.extend(cleaned_table)

        # === MENULIS KE EXCEL ===
        
        current_row = 1

        # A. Tulis Header (Judul Dokumen)
        # Style Bold untuk judul agar lebih rapi (Opsional, tapi bagus)
        from openpyxl.styles import Font
        bold_font = Font(bold=True)

        for line in header_text:
            cell = ws.cell(row=current_row, column=1, value=line)
            cell.font = bold_font
            current_row += 1
        
        current_row += 1 # Jarak 1 baris kosong

        # B. Tulis Data Tabel
        if all_table_data:
            df = pd.DataFrame(all_table_data)
            for r in dataframe_to_rows(df, index=False, header=False):
                ws.append(r)
        else:
            ws.cell(row=current_row, column=1, value="Tidak ada data tabel ditemukan.")

        wb.save(tmp_xlsx_path)

        background_tasks.add_task(cleanup_folder, tmp_dir)

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

# === FITUR 3: PDF KE PPTX (TEXT + IMAGES) ===
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

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        doc = fitz.open(tmp_pdf_path)

        for page_num, page in enumerate(doc):
            blank_slide_layout = prs.slide_layouts[6] 
            slide = prs.slides.add_slide(blank_slide_layout)

            # Gunakan get_text("dict") untuk mendapatkan struktur lebih lengkap (termasuk gambar)
            blocks = page.get_text("dict")["blocks"]
            
            for b in blocks:
                # Tipe 0 = Teks
                if b["type"] == 0: 
                    bbox = b["bbox"]
                    # Koordinat PDF
                    x0, y0, x1, y1 = bbox
                    
                    # Konversi ke PPTX (Inches)
                    left = Inches(x0 / 72)
                    top = Inches(y0 / 72)
                    width = Inches((x1 - x0) / 72)
                    height = Inches((y1 - y0) / 72)

                    # Gabungkan teks dalam satu blok
                    text_content = ""
                    for line in b["lines"]:
                        for span in line["spans"]:
                            text_content += span["text"] + " "
                    
                    if text_content.strip():
                        txBox = slide.shapes.add_textbox(left, top, width, height)
                        tf = txBox.text_frame
                        tf.word_wrap = True # Agar teks tidak memanjang ke samping
                        p = tf.add_paragraph()
                        p.text = text_content
                        p.font.size = Pt(11) # Ukuran font standar

                # Tipe 1 = Gambar
                elif b["type"] == 1:
                    bbox = b["bbox"]
                    x0, y0, x1, y1 = bbox
                    
                    left = Inches(x0 / 72)
                    top = Inches(y0 / 72)
                    width = Inches((x1 - x0) / 72)
                    height = Inches((y1 - y0) / 72)
                    
                    # Ekstrak data gambar
                    image_bytes = b["image"]
                    image_ext = b["ext"]
                    image_filename = os.path.join(tmp_dir, f"img_{page_num}_{id(b)}.{image_ext}")
                    
                    with open(image_filename, "wb") as img_file:
                        img_file.write(image_bytes)
                    
                    # Tempel gambar ke slide
                    try:
                        slide.shapes.add_picture(image_filename, left, top, width=width, height=height)
                    except Exception as img_err:
                        logging.warning(f"Gagal menambah gambar: {img_err}")

        prs.save(tmp_ppt_path)
        doc.close()

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
