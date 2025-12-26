# -*- coding: utf-8 -*-
import os
import shutil
import logging
import tempfile
import traceback
import io  # Penting untuk in-memory processing

# Framework
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

# Library Konversi
from pdf2docx import Converter
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

# Inisialisasi Aplikasi
app = FastAPI(
    title="Aplikasi Konverter PDF Pro",
    version="4.0 (VPS Optimized)",
)

# Setup CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # Ganti dengan domain frontend Anda untuk keamanan ekstra
    allow_methods=["*"],
    allow_headers=["*"],
)

# Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Konfigurasi
MAX_FILE_SIZE = 25 * 1024 * 1024  # Naikkan ke 25 MB karena server lebih kuat

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
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")
    
    file.file.seek(0, 2)
    file_size = file.file.tell()
    file.file.seek(0)
    
    if file_size > MAX_FILE_SIZE:
        raise HTTPException(status_code=400, detail=f"Ukuran file terlalu besar (Maks {MAX_FILE_SIZE/1024/1024}MB)")

@app.get("/")
def read_root():
    return {"message": "Server PDF Backend (DigitalOcean Optimized) is Running!"}

# === FITUR 1: PDF KE DOCX (STABIL) ===
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

        # pdf2docx sudah cukup optimal, kita biarkan default CPU count
        # agar tidak crash di dalam container Docker.
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
        logging.error("!!! ERROR PDF TO DOCX !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Word: {str(e)}")


# === FITUR 2: PDF KE EXCEL (SMART HEADER DETECTION) ===
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

        wb = Workbook()
        ws = wb.active
        ws.title = "Data Hasil Convert"

        all_table_data = []
        header_text = []

        with pdfplumber.open(tmp_pdf_path) as pdf:
            # 1. DETEKSI HEADER PINTAR (Hanya Halaman 1)
            if len(pdf.pages) > 0:
                first_page = pdf.pages[0]
                found_tables = first_page.find_tables()
                
                # Default: Ambil 15% atas jika tidak ketemu tabel
                crop_bottom = first_page.height * 0.15 
                
                if found_tables:
                    # Jika ada tabel, ambil area DI ATAS tabel tersebut
                    # bbox[1] adalah posisi Y (atas) dari tabel
                    crop_bottom = max(0, found_tables[0].bbox[1] - 5)
                
                # Crop dan Ambil Teks Header
                try:
                    header_crop = first_page.crop((0, 0, first_page.width, crop_bottom))
                    raw_header = header_crop.extract_text()
                    if raw_header:
                        header_text = raw_header.split('\n')
                except Exception:
                    # Fallback jika crop gagal
                    pass

            # 2. EKSTRAKSI TABEL (SEMUA HALAMAN DI-MERGE)
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    # Bersihkan None -> String kosong
                    cleaned_table = [[cell if cell is not None else "" for cell in row] for row in table]
                    all_table_data.extend(cleaned_table)

        # === PENULISAN KE EXCEL ===
        current_row = 1

        # Tulis Judul (Bold)
        bold_font = Font(bold=True)
        for line in header_text:
            cell = ws.cell(row=current_row, column=1, value=line)
            cell.font = bold_font
            current_row += 1
        
        current_row += 1 # Spasi

        # Tulis Tabel
        if all_table_data:
            df = pd.DataFrame(all_table_data)
            for r in dataframe_to_rows(df, index=False, header=False):
                ws.append(r)
        else:
            ws.cell(row=current_row, column=1, value="Tidak ada tabel terdeteksi.")

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


# === FITUR 3: PDF KE PPTX (HIGH SPEED IN-MEMORY) ===
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
        prs.slide_width = Inches(13.333) # Format Widescreen 16:9
        prs.slide_height = Inches(7.5)
        
        doc = fitz.open(tmp_pdf_path)

        for page in doc:
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # Layout Kosong

            # Gunakan get_text("dict") untuk parsing struktur lengkap
            blocks = page.get_text("dict")["blocks"]
            
            for b in blocks:
                # --- MENANGANI TEKS (EDITABLE) ---
                if b["type"] == 0: 
                    bbox = b["bbox"]
                    x0, y0, x1, y1 = bbox
                    
                    # Konversi Point PDF ke Inch PPT
                    left = Inches(x0 / 72)
                    top = Inches(y0 / 72)
                    width = Inches((x1 - x0) / 72)
                    height = Inches((y1 - y0) / 72)

                    text_content = ""
                    for line in b["lines"]:
                        for span in line["spans"]:
                            text_content += span["text"] + " "
                    
                    if text_content.strip():
                        # Buat Text Box
                        txBox = slide.shapes.add_textbox(left, top, width, height)
                        tf = txBox.text_frame
                        tf.word_wrap = True
                        p = tf.add_paragraph()
                        p.text = text_content
                        p.font.size = Pt(11) # Ukuran font aman

                # --- MENANGANI GAMBAR (IN-MEMORY FAST) ---
                elif b["type"] == 1:
                    bbox = b["bbox"]
                    x0, y0, x1, y1 = bbox
                    
                    left = Inches(x0 / 72)
                    top = Inches(y0 / 72)
                    width = Inches((x1 - x0) / 72)
                    height = Inches((y1 - y0) / 72)
                    
                    # INI KUNCI KECEPATAN:
                    # Langsung stream bytes gambar ke RAM, tidak perlu save ke file
                    image_bytes = b["image"]
                    image_stream = io.BytesIO(image_bytes)
                    
                    try:
                        slide.shapes.add_picture(image_stream, left, top, width=width, height=height)
                    except Exception:
                        # Skip jika gambar corrupt/format tidak didukung PPT
                        continue

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
