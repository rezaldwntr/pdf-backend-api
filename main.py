# -*- coding: utf-8 -*-
"""
Aplikasi Web Konverter PDF (c) 2024
Versi: 2.2 (Editable PPTX Conversion)
"""
import os
import shutil
import logging
import tempfile
import traceback
from zipfile import ZipFile

# Framework
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi import BackgroundTasks
from fastapi.responses import FileResponse

# Library Konversi
import camelot.io as camelot
from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Inches, Pt
from docx import Document
from docx.shared import Inches as DocxInches
from pdf2docx import Converter
import pandas as pd
import pdfplumber

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
    except: pass

# Konfigurasi logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# === ENDPOINTS ===

@app.get("/")
def read_root():
    return {"message": "Server Konverter PDF sedang berjalan."}

# === FITUR 1: PDF KE DOCX ===
@app.post("/convert/pdf-to-docx")
async def convert_pdf_to_docx(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    # Gunakan folder temp unik agar aman
    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    
    # Nama output docx
    docx_filename = os.path.splitext(file.filename)[0] + ".docx"
    tmp_docx_path = os.path.join(tmp_dir, docx_filename)

    try:
        # Simpan PDF
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Proses Konversi
        cv = Converter(tmp_pdf_path)
        cv.convert(tmp_docx_path, start=0, end=None, multiprocess=False, cpu_count=1)
        cv.close()

        # Jadwalkan Hapus Folder SETELAH file terkirim
        background_tasks.add_task(cleanup_folder, tmp_dir)

        # Kirim File
        return FileResponse(
            path=tmp_docx_path,
            filename=docx_filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        cleanup_folder(tmp_dir) # Hapus jika error
        print("!!! ERROR PDF TO DOCX !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Word: {str(e)}")

# === FITUR 2: PDF KE EXCEL ===
@app.post("/convert/pdf-to-excel")
async def convert_pdf_to_excel(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")
    
    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    excel_filename = os.path.splitext(file.filename)[0] + ".xlsx"
    tmp_excel_path = os.path.join(tmp_dir, excel_filename)

    try:
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Extract tables
        tables = camelot.read_pdf(tmp_pdf_path, pages='all', flavor='stream')
        
        # Save to Excel
        if len(tables) > 0:
            with pd.ExcelWriter(tmp_excel_path) as writer:
                for i, table in enumerate(tables):
                    sheet_name = f"Table {i+1}"
                    table.df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        else:
            # Empty excel if no tables
            pd.DataFrame(["No tables found"]).to_excel(tmp_excel_path, header=False, index=False)

        background_tasks.add_task(cleanup_folder, tmp_dir)
        return FileResponse(
            path=tmp_excel_path, 
            filename=excel_filename,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error(f"Error PDF to Excel: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Excel: {str(e)}")


# === FITUR 3: PDF KE PPTX (Editable - Text Extraction) ===
@app.post("/convert/pdf-to-ppt")
async def convert_pdf_to_ppt(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    ppt_filename = os.path.splitext(file.filename)[0] + ".pptx"
    tmp_ppt_path = os.path.join(tmp_dir, ppt_filename)

    try:
        # Simpan PDF
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        prs = Presentation()
        
        # Menggunakan pdfplumber untuk ekstraksi teks dan posisi yang lebih akurat
        with pdfplumber.open(tmp_pdf_path) as pdf:
            for page in pdf.pages:
                # Tambahkan slide kosong
                blank_slide_layout = prs.slide_layouts[6]
                slide = prs.slides.add_slide(blank_slide_layout)
                
                # Sesuaikan ukuran slide dengan halaman PDF (point to EMU)
                # 1 point = 12700 EMU
                width = int(page.width * 12700)
                height = int(page.height * 12700)
                prs.slide_width = width
                prs.slide_height = height

                # Ekstrak kata-kata/teks
                words = page.extract_words()
                
                # Mengelompokkan kata menjadi blok teks sederhana bisa menjadi kompleks.
                # Di sini kita akan mencoba membuat textbox untuk setiap kata atau grup kata
                # agar posisinya seakurat mungkin.
                # Untuk hasil yang lebih 'bersih', biasanya diperlukan algoritma pengelompokan yang lebih canggih.
                
                # Sederhananya, kita iterasi kata-kata dan menempatkannya.
                # Namun, menempatkan per kata bisa membuat PPT sangat berat dan sulit diedit.
                # Kita akan coba pendekatan per-baris jika memungkinkan atau menggunakan extract_text dengan layout=True.
                
                # Alternatif: Menggunakan layout clustering dari pdfplumber belum tentu sempurna.
                # Kita akan gunakan pendekatan iterasi 'words' untuk akurasi posisi, 
                # atau 'extract_text' untuk blok teks.
                
                # Mari gunakan pendekatan blok teks sederhana:
                # pdfplumber tidak memberikan blok teks secara langsung seperti PDFMiner, tapi kita bisa pakai extract_words
                # dan menggabungkan yang berdekatan.
                
                # Untuk kesederhanaan dan fungsionalitas edit:
                # Kita masukkan seluruh teks halaman ke satu textbox jika struktur tidak terlalu kompleks,
                # TAPI user minta "persis seperti aslinya".
                # Maka kita harus menempatkan elemen sesuai koordinatnya.
                
                for word in words:
                    x = Inches(float(word['x0']) / 72)
                    y = Inches(float(word['top']) / 72)
                    w = Inches((float(word['x1']) - float(word['x0'])) / 72)
                    h = Inches((float(word['bottom']) - float(word['top'])) / 72)
                    
                    textbox = slide.shapes.add_textbox(x, y, w, h)
                    tf = textbox.text_frame
                    tf.word_wrap = False # Agar tidak wrap aneh-aneh
                    p = tf.paragraphs[0]
                    p.text = word['text']
                    p.font.size = Pt(float(word.get('size', 12))) # Ukuran font
                    # Kita bisa menambahkan font name jika ada info, tapi pdfplumber terbatas soal ini.

        # Simpan PPTX
        prs.save(tmp_ppt_path)

        background_tasks.add_task(cleanup_folder, tmp_dir)

        return FileResponse(
            path=tmp_ppt_path,
            filename=ppt_filename,
            media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error(f"Error PDF to PPT: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert PPT: {str(e)}")

# === FITUR 4: PDF KE GAMBAR (ZIP) ===
@app.post("/convert/pdf-to-image")
async def convert_pdf_to_image(
    background_tasks: BackgroundTasks,
    output_format: str = "png", 
    file: UploadFile = File(...)
):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    tmp_dir = tempfile.mkdtemp()
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    zip_filename = os.path.splitext(file.filename)[0] + ".zip"
    tmp_zip_path = os.path.join(tmp_dir, zip_filename)

    try:
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        images = convert_from_path(tmp_pdf_path)
        
        with ZipFile(tmp_zip_path, 'w') as zipf:
            for i, img in enumerate(images):
                img_name = f"page_{i+1}.{output_format}"
                img_path = os.path.join(tmp_dir, img_name)
                img.save(img_path, output_format.upper())
                zipf.write(img_path, img_name)

        background_tasks.add_task(cleanup_folder, tmp_dir)
        return FileResponse(tmp_zip_path, filename=zip_filename, media_type='application/zip')

    except Exception as e:
        cleanup_folder(tmp_dir)
        logging.error(f"Error PDF to Image: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Image: {str(e)}")
