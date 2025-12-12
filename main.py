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


# === FITUR 2: PDF KE EXCEL (ADVANCED) ===
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
        footer_text = []

        with pdfplumber.open(tmp_pdf_path) as pdf:
            # 1. AMBIL HEADER (JUDUL) DARI HALAMAN PERTAMA
            # Asumsi: Judul ada di bagian atas halaman 1 (misal 20% teratas)
            first_page = pdf.pages[0]
            # Ambil teks area atas (header)
            # Crop box: (x0, top, x1, bottom)
            header_crop = first_page.crop((0, 0, first_page.width, first_page.height * 0.2)) 
            raw_header = header_crop.extract_text()
            if raw_header:
                header_text = raw_header.split('\n')

            # 2. AMBIL TABEL DARI SEMUA HALAMAN (MERGE SHEET)
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    # Bersihkan data None/Null
                    cleaned_table = [[cell if cell is not None else "" for cell in row] for row in table]
                    all_table_data.extend(cleaned_table)

            # 3. AMBIL FOOTER (TANDA TANGAN) DARI HALAMAN TERAKHIR
            # Asumsi: Tanda tangan ada di 20% terbawah halaman terakhir
            last_page = pdf.pages[-1]
            footer_crop = last_page.crop((0, last_page.height * 0.8, last_page.width, last_page.height))
            raw_footer = footer_crop.extract_text()
            if raw_footer:
                footer_text = raw_footer.split('\n')

        # === MENULIS KE EXCEL ===
        
        # A. Tulis Header (Judul)
        current_row = 1
        for line in header_text:
            ws.cell(row=current_row, column=1, value=line)
            current_row += 1
        
        current_row += 1 # Kasih jarak 1 baris kosong

        # B. Tulis Data Tabel (Merged)
        if all_table_data:
            # Menggunakan pandas agar lebih rapi handling datanya, lalu convert ke rows openpyxl
            df = pd.DataFrame(all_table_data)
            
            # Jika baris pertama PDF terdeteksi sebagai header berulang, kita bisa skip di logic ini
            # Tapi untuk aman, kita dump semua dulu.
            
            for r in dataframe_to_rows(df, index=False, header=False):
                ws.append(r)
                current_row += 1
        else:
            ws.cell(row=current_row, column=1, value="Tidak ada data tabel ditemukan.")
            current_row += 1

        current_row += 2 # Jarak sebelum tanda tangan

        # C. Tulis Footer (Tanda Tangan)
        # Tantangan: Membuat Zig-Zag (1... 2...) otomatis itu sulit tanpa koordinat pasti.
        # Kita akan dump teksnya, user tinggal geser kolomnya.
        ws.cell(row=current_row, column=1, value="--- Bagian Tanda Tangan ---")
        current_row += 1
        for line in footer_text:
            # Coba deteksi jika ada pola "Nama ...... (Jarak) ...... Nama"
            # Split berdasarkan spasi lebar
            parts = line.split("   ") 
            col_idx = 1
            for part in parts:
                if part.strip():
                    ws.cell(row=current_row, column=col_idx, value=part.strip())
                    # Lompat kolom biar ada efek jarak (simulasi zig-zag sederhana)
                    col_idx += 3 
            current_row += 1

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
