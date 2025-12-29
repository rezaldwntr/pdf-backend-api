# -*- coding: utf-8 -*-
"""
Aplikasi Web Konverter PDF (c) 2024
Versi: 2.6 (Full Page Layout: Text + Tables + Smart Alignment)
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
    version="2.6",
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
    return {"message": "Server PDF Backend (Full Layout Engine) is Running!"}

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

# === FITUR 2: PDF KE EXCEL (FULL LAYOUT: TEXT + TABLE) ===
# Helper function: Cek apakah kata ada di dalam area tabel
def is_inside_table(word_bbox, table_bboxes):
    # word_bbox = (x0, top, x1, bottom)
    # table_bbox = (x0, top, x1, bottom)
    wx0, wtop, wx1, wbottom = word_bbox
    w_mid_x = (wx0 + wx1) / 2
    w_mid_y = (wtop + wbottom) / 2
    
    for tbox in table_bboxes:
        tx0, ttop, tx1, tbottom = tbox
        # Cek apakah titik tengah kata ada di dalam kotak tabel
        if tx0 <= w_mid_x <= tx1 and ttop <= w_mid_y <= tbottom:
            return True
    return False

# Helper function: Gabungkan kata-kata menjadi baris kalimat
def cluster_words_into_lines(words, tolerance=3):
    lines = []
    if not words:
        return lines
        
    # Urutkan berdasarkan posisi Y (top)
    sorted_words = sorted(words, key=lambda w: w['top'])
    
    current_line = [sorted_words[0]]
    
    for i in range(1, len(sorted_words)):
        word = sorted_words[i]
        prev_word = current_line[-1]
        
        # Jika posisi Y berdekatan (toleransi piksel), anggap satu baris
        if abs(word['top'] - prev_word['top']) < tolerance:
            current_line.append(word)
        else:
            lines.append(current_line)
            current_line = [word]
    lines.append(current_line)
    
    # Proses setiap line untuk menggabungkan teks
    final_lines = []
    for line in lines:
        # Urutkan kata dari kiri ke kanan (x0)
        line.sort(key=lambda w: w['x0'])
        text_content = " ".join([w['text'] for w in line])
        
        # Hitung bounding box baris ini
        x0 = min(w['x0'] for w in line)
        top = min(w['top'] for w in line)
        x1 = max(w['x1'] for w in line)
        bottom = max(w['bottom'] for w in line)
        
        final_lines.append({
            'text': text_content,
            'bbox': (x0, top, x1, bottom),
            'top': top,
            'x0': x0,
            'x1': x1
        })
        
    return final_lines

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
                
                current_excel_row = 1
                
                # Style Definitions
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                header_font = Font(bold=True)
                header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

                for page_idx, page in enumerate(pdf.pages):
                    page_width = page.width
                    
                    # 1. AMBIL TABEL
                    tables = page.find_tables()
                    table_bboxes = [t.bbox for t in tables] # (x0, top, x1, bottom)
                    
                    # 2. AMBIL TEXT (NON-TABEL)
                    words = page.extract_words()
                    non_table_words = []
                    
                    for w in words:
                        # Format w: {'text': '...', 'x0': 10, 'top': 20, ...}
                        # Ubah ke tuple bbox untuk pengecekan
                        w_bbox = (w['x0'], w['top'], w['x1'], w['bottom'])
                        if not is_inside_table(w_bbox, table_bboxes):
                            non_table_words.append(w)
                    
                    # Group words menjadi baris kalimat
                    text_lines = cluster_words_into_lines(non_table_words)
                    
                    # 3. GABUNGKAN SEMUA ELEMENT (TABEL + TEKS)
                    # Kita buat list "master" yang berisi semua objek
                    # tipe: 'text' atau 'table'
                    page_elements = []
                    
                    for t in tables:
                        page_elements.append({
                            'type': 'table',
                            'top': t.bbox[1], # Posisi Y atas
                            'obj': t
                        })
                        
                    for l in text_lines:
                        page_elements.append({
                            'type': 'text',
                            'top': l['top'],
                            'obj': l
                        })
                    
                    # 4. URUTKAN BERDASARKAN POSISI Y (ATAS KE BAWAH)
                    page_elements.sort(key=lambda x: x['top'])
                    
                    # 5. TULIS KE EXCEL
                    for element in page_elements:
                        if element['type'] == 'text':
                            line = element['obj']
                            text = line['text']
                            x0 = line['x0']
                            x1 = line['x1']
                            
                            cell = worksheet.cell(row=current_excel_row, column=1, value=text)
                            
                            # Logika Alignment Sederhana
                            # Hitung titik tengah teks
                            text_center = (x0 + x1) / 2
                            page_center = page_width / 2
                            
                            # Jika tengah teks berada di dekat tengah halaman (toleransi 10%)
                            if abs(text_center - page_center) < (page_width * 0.1):
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                                # Merge cells agar rapi (misal kolom A-H)
                                worksheet.merge_cells(start_row=current_excel_row, start_column=1, end_row=current_excel_row, end_column=8)
                            elif x0 > (page_width * 0.6): # Jika mulai di sebelah kanan
                                cell.alignment = Alignment(horizontal='right')
                            else:
                                cell.alignment = Alignment(horizontal='left')
                                
                            # Jika teks tebal/bold (simulasi: biasanya judul di atas itu bold)
                            # Karena pdfplumber extract_words tidak return font weight secara akurat, 
                            # kita asumsikan baris sendiri di luar tabel mungkin judul jika pendek.
                            
                            current_excel_row += 1
                        
                        elif element['type'] == 'table':
                            table = element['obj']
                            data = table.extract()
                            if not data: continue
                            
                            # Bersihkan data
                            clean_data = [[c if c is not None else "" for c in r] for r in data]
                            df = pd.DataFrame(clean_data)
                            
                            # Tulis tabel
                            start_row = current_excel_row
                            df.to_excel(writer, sheet_name="Hasil Konversi", startrow=start_row-1, startcol=0, index=False, header=False)
                            
                            # Styling Tabel
                            end_row = start_row + len(df)
                            end_col = len(df.columns)
                            
                            # Cek alignment header tabel asli (row pertama)
                            header_rects = table.rows[0]
                            
                            for r_idx, row_cells in enumerate(worksheet.iter_rows(min_row=start_row, max_row=end_row-1, min_col=1, max_col=end_col)):
                                is_header = (r_idx == 0)
                                for c_idx, cell in enumerate(row_cells):
                                    cell.border = thin_border
                                    cell.alignment = Alignment(vertical='top', wrap_text=True)
                                    
                                    if is_header:
                                        cell.font = header_font
                                        cell.fill = header_fill
                                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                                        # (Logika alignment spesifik per kolom bisa ditambahkan disini seperti versi sebelumnya jika perlu)
                            
                            # Update pointer baris
                            current_excel_row = end_row + 1 # Beri jarak 1 baris
                    
                    current_excel_row += 2 # Jarak antar halaman PDF

                # Finishing: Auto Width Column
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if cell.value and len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except: pass
                    adjusted_width = (max_length + 2)
                    if adjusted_width > 60: adjusted_width = 60
                    worksheet.column_dimensions[column].width = adjusted_width

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
