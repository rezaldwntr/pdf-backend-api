# -*- coding: utf-8 -*-
"""
Aplikasi Web Konverter PDF (c) 2024
Versi: 2.4 (Excel Styling & Layout Fix)
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
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# Konfigurasi
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10 MB

# Inisialisasi Aplikasi
app = FastAPI(
    title="Aplikasi Konverter PDF",
    description="API untuk mengubah format file dari PDF ke format lainnya.",
    version="2.4",
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
    return {"message": "Server PDF Backend (Excel Enhanced) is Running!"}

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

# === FITUR 2: PDF KE EXCEL (DIPERBAIKI: Styling & Layout) ===
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

        # Menggunakan OpenPyXL Engine agar bisa styling
        with pd.ExcelWriter(tmp_xlsx_path, engine='openpyxl') as writer:
            with pdfplumber.open(tmp_pdf_path) as pdf:
                
                # Kita akan menulis manual ke worksheet agar bisa kontrol layout
                # Inisialisasi workbook dan sheet
                writer.book.create_sheet("Hasil Konversi")
                worksheet = writer.book["Hasil Konversi"]
                
                # Hapus sheet default "Sheet" jika ada
                if "Sheet" in writer.book.sheetnames:
                    del writer.book["Sheet"]
                
                current_row = 1 # Pointer baris Excel
                has_data = False

                # Style Border Tipis
                thin_border = Border(left=Side(style='thin'), 
                                     right=Side(style='thin'), 
                                     top=Side(style='thin'), 
                                     bottom=Side(style='thin'))

                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    
                    for table in tables:
                        if not table:
                            continue
                            
                        has_data = True
                        
                        # Bersihkan data (None -> "")
                        clean_table = [[cell if cell is not None else "" for cell in row] for row in table]
                        df = pd.DataFrame(clean_table)
                        
                        # 1. Tulis Judul/Header Kecil penanda Halaman (Opsional, agar rapi)
                        header_cell = worksheet.cell(row=current_row, column=1, value=f"Page {i+1} - Table")
                        header_cell.font = Font(bold=True, color="0000FF")
                        current_row += 1
                        
                        # 2. Tulis Tabel ke Excel di posisi 'current_row'
                        # index=False, header=False agar murni data tabel PDF
                        df.to_excel(writer, sheet_name="Hasil Konversi", startrow=current_row-1, startcol=0, index=False, header=False)
                        
                        # 3. APPLY STYLING (Borders & Alignment)
                        # Loop area yang baru saja ditulis untuk memberi garis
                        end_row = current_row + len(df)
                        end_col = len(df.columns)
                        
                        for r in range(current_row, end_row):
                            for c in range(1, end_col + 1):
                                cell = worksheet.cell(row=r, column=c)
                                cell.border = thin_border
                                cell.alignment = Alignment(wrap_text=True, vertical='top')
                        
                        # Update pointer baris (beri jarak 2 baris antar tabel)
                        current_row = end_row + 2

                # 4. Auto-adjust Column Width (Finishing Touch)
                # Agar lebar kolom menyesuaikan isi teks
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter # Get the column name
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    # Set lebar (limit max 50 agar tidak terlalu lebar)
                    adjusted_width = (max_length + 2)
                    if adjusted_width > 50: adjusted_width = 50
                    worksheet.column_dimensions[column].width = adjusted_width

                if not has_data:
                    worksheet.cell(row=1, column=1, value="Tidak ada tabel yang terdeteksi dalam PDF ini.")

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
