# -*- coding: utf-8 -*-
"""
Aplikasi Web Konverter PDF (c) 2024
Versi: 5.0 (PPTX Ultimate Precision: Zero-Margin Text, Color Accuracy & Adaptive Layout)
"""
import os
import shutil
import logging
import tempfile
import io
from zipfile import ZipFile

# Framework
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

# Library Konversi
from pdf2docx import Converter
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
import pdfplumber
import pandas as pd
from PIL import Image # Helper untuk cek gambar

# Library Excel Styling
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Konfigurasi Ukuran File (25MB)
MAX_FILE_SIZE = 25 * 1024 * 1024 

# Inisialisasi Aplikasi
app = FastAPI(
    title="Aplikasi Konverter PDF",
    description="API untuk mengubah format file dari PDF ke format lainnya.",
    version="5.0",
)

# === KONFIGURASI CORS ===
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
    return {"message": "Server PDF Backend (V5.0 Ultimate Precision) is Running!"}

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

        # Konversi PDF ke DOCX (Multiprocess False wajib untuk Docker)
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

# === FITUR 2: PDF KE EXCEL ===
def is_inside_table(word_bbox, table_bboxes):
    wx0, wtop, wx1, wbottom = word_bbox
    w_mid_x = (wx0 + wx1) / 2
    w_mid_y = (wtop + wbottom) / 2
    for tbox in table_bboxes:
        tx0, ttop, tx1, tbottom = tbox
        if tx0 <= w_mid_x <= tx1 and ttop <= w_mid_y <= tbottom:
            return True
    return False

def cluster_words_into_lines(words, tolerance=3):
    lines = []
    if not words: return lines
    sorted_words = sorted(words, key=lambda w: w['top'])
    current_line = [sorted_words[0]]
    for i in range(1, len(sorted_words)):
        word = sorted_words[i]
        prev_word = current_line[-1]
        if abs(word['top'] - prev_word['top']) < tolerance:
            current_line.append(word)
        else:
            lines.append(current_line)
            current_line = [word]
    lines.append(current_line)
    final_lines = []
    for line in lines:
        line.sort(key=lambda w: w['x0'])
        text_content = " ".join([w['text'] for w in line])
        x0 = min(w['x0'] for w in line)
        top = min(w['top'] for w in line)
        x1 = max(w['x1'] for w in line)
        bottom = max(w['bottom'] for w in line)
        final_lines.append({'text': text_content, 'bbox': (x0, top, x1, bottom), 'top': top, 'x0': x0, 'x1': x1})
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

        with pd.ExcelWriter(tmp_xlsx_path, engine='openpyxl') as writer:
            with pdfplumber.open(tmp_pdf_path) as pdf:
                writer.book.create_sheet("Hasil Konversi")
                worksheet = writer.book["Hasil Konversi"]
                if "Sheet" in writer.book.sheetnames: del writer.book["Sheet"]
                
                current_excel_row = 1
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                header_font = Font(bold=True)
                header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

                for page in pdf.pages:
                    page_width = page.width
                    tables = page.find_tables()
                    table_bboxes = [t.bbox for t in tables]
                    words = page.extract_words()
                    non_table_words = [w for w in words if not is_inside_table((w['x0'], w['top'], w['x1'], w['bottom']), table_bboxes)]
                    text_lines = cluster_words_into_lines(non_table_words)
                    
                    page_elements = []
                    for t in tables:
                        try:
                            t_data = t.extract()
                            col_count = len(t_data[0]) if t_data and len(t_data) > 0 else 1
                        except: col_count = 1
                        page_elements.append({'type': 'table', 'top': t.bbox[1], 'obj': t, 'cols': col_count})
                    for l in text_lines:
                        page_elements.append({'type': 'text', 'top': l['top'], 'obj': l})
                    
                    page_elements.sort(key=lambda x: x['top'])
                    
                    for i, element in enumerate(page_elements):
                        if element['type'] == 'text':
                            line = element['obj']
                            cell = worksheet.cell(row=current_excel_row, column=1, value=line['text'])
                            text_center = (line['x0'] + line['x1']) / 2
                            if abs(text_center - (page_width / 2)) < (page_width * 0.1):
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                                target_merge_col = 1
                                for next_el in page_elements[i+1:]:
                                    if next_el['type'] == 'table':
                                        target_merge_col = next_el['cols']
                                        break
                                if target_merge_col > 1:
                                    worksheet.merge_cells(start_row=current_excel_row, start_column=1, end_row=current_excel_row, end_column=target_merge_col)
                            elif line['x0'] > (page_width * 0.6):
                                cell.alignment = Alignment(horizontal='right')
                            else:
                                cell.alignment = Alignment(horizontal='left')
                            current_excel_row += 1
                        elif element['type'] == 'table':
                            data = element['obj'].extract()
                            if not data: continue
                            clean_data = [[c if c is not None else "" for c in r] for r in data]
                            df = pd.DataFrame(clean_data)
                            start_row = current_excel_row
                            df.to_excel(writer, sheet_name="Hasil Konversi", startrow=start_row-1, startcol=0, index=False, header=False)
                            end_row = start_row + len(df)
                            end_col = len(df.columns)
                            for r_idx, row_cells in enumerate(worksheet.iter_rows(min_row=start_row, max_row=end_row-1, min_col=1, max_col=end_col)):
                                is_header = (r_idx == 0)
                                for cell in row_cells:
                                    cell.border = thin_border
                                    cell.alignment = Alignment(vertical='top', wrap_text=True)
                                    if is_header:
                                        cell.font = header_font
                                        cell.fill = header_fill
                                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            current_excel_row = end_row + 1
                    current_excel_row += 2

                if worksheet.max_column:
                    for col_idx in range(1, worksheet.max_column + 1):
                        col_letter = get_column_letter(col_idx)
                        max_len = 0
                        for row_idx in range(1, worksheet.max_row + 1):
                            try:
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                if cell.value: max_len = max(max_len, len(str(cell.value)))
                            except: pass
                        worksheet.column_dimensions[col_letter].width = min(max_len + 2, 60)

        background_tasks.add_task(cleanup_folder, tmp_dir)
        return FileResponse(path=tmp_xlsx_path, filename=xlsx_filename, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error(f"ERROR PDF TO EXCEL: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Gagal convert Excel: {str(e)}")

# === FITUR 3: PDF KE PPTX (ULTIMATE PRECISION V5) ===
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
        # Set slide size default ke widescreen dulu, nanti diupdate per page jika perlu
        # Tapi python-pptx set size itu global. Jadi kita ambil size halaman pertama PDF.
        doc = fitz.open(tmp_pdf_path)
        
        if len(doc) > 0:
            p1 = doc[0]
            # Convert points to EMUs (1 inch = 72 pt = 914400 EMUs)
            # PyMuPDF: pt. python-pptx: EMU/Inches.
            # Inches(x) -> x * 914400. 
            # 1 pt = 1/72 inch.
            width_pt = p1.rect.width
            height_pt = p1.rect.height
            prs.slide_width = int((width_pt / 72) * 914400)
            prs.slide_height = int((height_pt / 72) * 914400)

        for page_index, page in enumerate(doc):
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # Layout 6 = Blank
            
            # --- ALGORITMA 1: GAMBAR & VEKTOR (Extract Raw) ---
            # Kita ambil list gambar dari halaman
            image_list = page.get_images(full=True)
            
            # Agar posisi gambar akurat, kita tidak pakai get_images langsung,
            # tapi via get_text("dict") yang memuat posisi (bbox) gambar.
            
            # Gunakan flags lengkap untuk ekstraksi maksimal
            flags = fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_PRESERVE_IMAGES
            blocks = page.get_text("dict", flags=flags)["blocks"]
            
            for b in blocks:
                # TYPE 1: GAMBAR (RASTER IMAGES)
                if b['type'] == 1:
                    img_bytes = b["image"]
                    ext = b["ext"]
                    # Posisi di PDF
                    x0, y0, x1, y1 = b["bbox"]
                    
                    # Simpan sementara
                    img_name = f"p{page_index}_img.{ext}"
                    img_path = os.path.join(tmp_dir, img_name)
                    with open(img_path, "wb") as f:
                        f.write(img_bytes)
                    
                    # Tempel ke PPT
                    try:
                        slide.shapes.add_picture(
                            img_path, 
                            Inches(x0 / 72), Inches(y0 / 72), 
                            width=Inches((x1 - x0) / 72), 
                            height=Inches((y1 - y0) / 72)
                        )
                    except Exception as img_err:
                        logging.error(f"Failed adding image: {img_err}")

                # TYPE 0: TEKS (EDITABLE)
                elif b['type'] == 0:
                    for line in b["lines"]:
                        # Kita buat Text Box PER BARIS agar posisi vertikal akurat
                        # Menggunakan bbox baris, bukan bbox blok
                        lx0, ly0, lx1, ly1 = line["bbox"]
                        
                        # Hitung lebar & tinggi
                        w = lx1 - lx0
                        h = ly1 - ly0
                        
                        # Buat Text Box
                        txBox = slide.shapes.add_textbox(
                            Inches(lx0 / 72), Inches(ly0 / 72), 
                            Inches(w / 72), Inches(h / 72)
                        )
                        
                        # --- MODIFIKASI KRUSIAL UNTUK "RAPI" ---
                        tf = txBox.text_frame
                        tf.word_wrap = False  # Matikan word wrap agar tidak turun baris sendiri
                        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT # Sesuaikan box dengan teks
                        
                        # SET MARGIN 0 (Kunci agar teks tidak bergeser)
                        tf.margin_left = 0
                        tf.margin_right = 0
                        tf.margin_top = 0
                        tf.margin_bottom = 0
                        tf.vertical_anchor = MSO_ANCHOR.TOP # Anchor ke atas
                        
                        p = tf.paragraphs[0]
                        
                        # Isi teks per kata/style (Span)
                        for span in line["spans"]:
                            text = span["text"]
                            # Filter karakter aneh
                            if not text.strip(): 
                                if text == " ": pass # keep spaces
                                else: continue

                            run = p.add_run()
                            run.text = text
                            
                            # 1. Font Size
                            run.font.size = Pt(span["size"])
                            
                            # 2. Font Name (Optional, fallback ke default PPT jika font PDF aneh)
                            # run.font.name = span["font"] 
                            
                            # 3. Warna (Color)
                            try:
                                c = span["color"]
                                # PyMuPDF sRGB integer conversion
                                r = (c >> 16) & 0xFF
                                g = (c >> 8) & 0xFF
                                b = c & 0xFF
                                run.font.color.rgb = RGBColor(r, g, b)
                            except:
                                run.font.color.rgb = RGBColor(0, 0, 0) # Default Hitam
                            
                            # 4. Styling (Bold/Italic)
                            flags_font = span["flags"]
                            # Cek dokumentasi PyMuPDF flags: 2^4 = Bold, 2^1 = Italic
                            if flags_font & 16:
                                run.font.bold = True
                            if flags_font & 2:
                                run.font.italic = True

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
            for i, page in enumerate(doc):
                pix = page.get_pixmap(dpi=200)
                img_name = f"page_{i+1}.{output_format}"
                img_path = os.path.join(tmp_dir, img_name)
                pix.save(img_path, output="jpg" if output_format.lower() in ['jpg','jpeg'] else output_format)
                zipf.write(img_path, img_name)
        
        doc.close()
        background_tasks.add_task(cleanup_folder, tmp_dir)
        return FileResponse(path=tmp_zip_path, filename=zip_filename, media_type='application/zip')

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error(f"ERROR PDF TO IMAGE: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Gagal convert Image: {str(e)}")
