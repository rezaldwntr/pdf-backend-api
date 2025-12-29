# -*- coding: utf-8 -*-
"""
Aplikasi Web Konverter PDF (c) 2024
Versi: 2.5 (Excel Header Auto-Align & Styling)
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
import pdfplumber
import pandas as pd

# Library Excel Styling
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Konfigurasi
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10 MB

# Inisialisasi Aplikasi
app = FastAPI(
    title="Aplikasi Konverter PDF",
    description="API untuk mengubah format file dari PDF ke format lainnya.",
    version="2.5",
)

# Mengizinkan akses dari Frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Helper untuk hapus folder
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
    return {"message": "Server PDF Backend (Excel Smart Align) is Running!"}

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

# === FITUR 2: PDF KE EXCEL (DIPERBAIKI: Header & Smart Alignment) ===
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

        # Gunakan Engine OpenPyXL
        with pd.ExcelWriter(tmp_xlsx_path, engine='openpyxl') as writer:
            with pdfplumber.open(tmp_pdf_path) as pdf:
                
                writer.book.create_sheet("Hasil Konversi")
                worksheet = writer.book["Hasil Konversi"]
                
                if "Sheet" in writer.book.sheetnames:
                    del writer.book["Sheet"]
                
                current_row = 1
                has_data = False

                # Style Setup
                thin_border = Border(left=Side(style='thin'), 
                                     right=Side(style='thin'), 
                                     top=Side(style='thin'), 
                                     bottom=Side(style='thin'))
                
                header_font = Font(bold=True, color="000000")
                header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid") # Abu-abu muda

                for i, page in enumerate(pdf.pages):
                    # Gunakan find_tables agar kita dapat koordinat sel (untuk deteksi alignment)
                    tables = page.find_tables()
                    
                    for table in tables:
                        if not table: continue
                        
                        data = table.extract()
                        if not data: continue
                        
                        has_data = True
                        
                        # 1. Judul Penanda Tabel
                        title_cell = worksheet.cell(row=current_row, column=1, value=f"Tabel (Halaman {i+1})")
                        title_cell.font = Font(bold=True, color="0000FF")
                        current_row += 1
                        
                        # 2. Tulis Data ke Excel
                        clean_data = [[c if c is not None else "" for c in r] for r in data]
                        df = pd.DataFrame(clean_data)
                        
                        start_row_idx = current_row
                        # Write data (header=False, karena kita handle manual stylenya)
                        df.to_excel(writer, sheet_name="Hasil Konversi", startrow=current_row-1, startcol=0, index=False, header=False)
                        
                        # 3. LOGIKA STYLING & ALIGNMENT
                        # Ambil geometri sel baris pertama (Header) dari objek table pdfplumber
                        # table.rows[0] berisi list bbox: (x0, top, x1, bottom)
                        header_rects = table.rows[0]

                        end_row = current_row + len(df)
                        end_col = len(df.columns)
                        
                        # Loop setiap sel di Excel yang baru ditulis
                        for r_idx, row_cells in enumerate(worksheet.iter_rows(min_row=start_row_idx, max_row=end_row-1, min_col=1, max_col=end_col)):
                            is_header = (r_idx == 0) # Apakah ini baris pertama?
                            
                            for c_idx, cell in enumerate(row_cells):
                                # Default Styles
                                cell.border = thin_border
                                horz_align = 'left' # Default alignment
                                vert_align = 'top'
                                
                                # --- HEADER HANDLING ---
                                if is_header:
                                    cell.font = header_font
                                    cell.fill = header_fill
                                    vert_align = 'center'
                                    
                                    # Deteksi Posisi Asli (Center/Left/Right)
                                    try:
                                        # Pastikan kolom ada di data geometri PDF
                                        if c_idx < len(header_rects):
                                            cell_rect = header_rects[c_idx]
                                            if cell_rect:
                                                # Crop PDF ke area sel spesifik ini
                                                # Tambah padding sedikit (+1/-1) agar text tepi tidak terpotong
                                                bbox = (cell_rect[0]-2, cell_rect[1]-2, cell_rect[2]+2, cell_rect[3]+2)
                                                cropped_cell = page.crop(bbox)
                                                
                                                # Ekstrak kata beserta koordinatnya
                                                words = cropped_cell.extract_words()
                                                
                                                if words:
                                                    # Hitung pusat teks vs pusat sel
                                                    w_x0 = min(w['x0'] for w in words) # Kiri Teks
                                                    w_x1 = max(w['x1'] for w in words) # Kanan Teks
                                                    
                                                    text_center = (w_x0 + w_x1) / 2
                                                    cell_center = (cell_rect[0] + cell_rect[2]) / 2
                                                    cell_width = cell_rect[2] - cell_rect[0]
                                                    
                                                    # LOGIKA DETEKSI:
                                                    # 1. Jika selisih tengah teks & tengah sel < 15% lebar sel -> CENTER
                                                    if abs(text_center - cell_center) < (cell_width * 0.15):
                                                        horz_align = 'center'
                                                    # 2. Jika jarak kanan lebih kecil dari jarak kiri -> RIGHT
                                                    elif (cell_rect[2] - w_x1) < (w_x0 - cell_rect[0]):
                                                        horz_align = 'right'
                                                    # 3. Sisanya -> LEFT
                                                    else:
                                                        horz_align = 'left'
                                    except Exception:
                                        pass # Fallback ke left jika deteksi gagal

                                cell.alignment = Alignment(horizontal=horz_align, vertical=vert_align, wrap_text=True)

                        current_row = end_row + 2

                # 4. Auto Width Columns (Agar teks tidak terpotong)
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if cell.value and len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    if adjusted_width > 60: adjusted_width = 60 # Limit max width
                    worksheet.column_dimensions[column].width = adjusted_width

                if not has_data:
                    worksheet.cell(row=1, column=1, value="Tidak ada tabel ditemukan.")

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

# === FITUR 3: PDF KE PPTX (Editable Text) ===
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
        doc = fitz.open(tmp_pdf_path)

        for page_num, page in enumerate(doc):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            blocks = page.get_text("dict")["blocks"]
            
            for b in blocks:
                if b['type'] == 0:
                    for line in b["lines"]:
                        for span in line["spans"]:
                            text = span["text"]
                            if not text.strip():
                                continue
                            x, y = span["bbox"][:2]
                            txBox = slide.shapes.add_textbox(Inches(x / 72), Inches(y / 72), Inches(5), Inches(0.5))
                            txBox.text_frame.text = text

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
                pix = page.get_pixmap(dpi=200)
                img_name = f"page_{page_num + 1}.{output_format}"
                img_path = os.path.join(tmp_dir, img_name)
                
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
