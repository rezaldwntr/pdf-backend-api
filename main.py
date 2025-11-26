import os
import shutil
import tempfile
import traceback
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
import io
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter
from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt

app = FastAPI()

# --- CORS ---
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def home():
    return {
        "status": "Server Online", 
        "endpoints": [
            "/convert/pdf-to-docx", 
            "/convert/pdf-to-excel",
            "/convert/pdf-to-ppt"
        ]
    }

# === FITUR 1: PDF KE WORD ===
@app.post("/convert/pdf-to-docx")
async def convert_pdf_to_docx(file: UploadFile = File(...)):
    # ... (KODE LAMA JANGAN DIUBAH) ...
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")
    tmp_pdf_path = None
    tmp_docx_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            shutil.copyfileobj(file.file, tmp_pdf)
            tmp_pdf_path = tmp_pdf.name
        tmp_docx_path = tmp_pdf_path.replace(".pdf", ".docx")
        original_filename = os.path.splitext(file.filename)[0] + ".docx"
        cv = Converter(tmp_pdf_path)
        cv.convert(tmp_docx_path, start=0, end=None, multiprocess=False, cpu_count=1)
        cv.close()
        return FileResponse(path=tmp_docx_path, filename=original_filename, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Word: {str(e)}")
    finally:
        if tmp_pdf_path and os.path.exists(tmp_pdf_path):
            try: os.remove(tmp_pdf_path)
            except: pass

# === FITUR 2: PDF KE EXCEL ===
@app.post("/convert/pdf-to-excel")
async def convert_pdf_to_excel(file: UploadFile = File(...)):
    # ... (KODE LAMA JANGAN DIUBAH) ...
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")
    tmp_pdf_path = None
    tmp_xlsx_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            shutil.copyfileobj(file.file, tmp_pdf)
            tmp_pdf_path = tmp_pdf.name
        tmp_xlsx_path = tmp_pdf_path.replace(".pdf", ".xlsx")
        original_filename = os.path.splitext(file.filename)[0] + ".xlsx"
        with pdfplumber.open(tmp_pdf_path) as pdf:
            with pd.ExcelWriter(tmp_xlsx_path, engine='openpyxl') as writer:
                tables_found = False
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    if tables:
                        tables_found = True
                        row_tracker = 0
                        for table in tables:
                            df = pd.DataFrame(table)
                            sheet_name = f"Page {i+1}"
                            df.to_excel(writer, sheet_name=sheet_name, startrow=row_tracker, index=False, header=False)
                            row_tracker += len(df) + 2
                if not tables_found:
                    df_fallback = pd.DataFrame(["Maaf, tidak ditemukan struktur tabel yang jelas."])
                    df_fallback.to_excel(writer, sheet_name="Info", index=False, header=False)
        return FileResponse(path=tmp_xlsx_path, filename=original_filename, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Excel: {str(e)}")
    finally:
        if tmp_pdf_path and os.path.exists(tmp_pdf_path):
            try: os.remove(tmp_pdf_path)
            except: pass


# === FITUR 3: PDF KE PPTX (VERSI STABIL / GAMBAR FULL) ===
@app.post("/convert/pdf-to-ppt")
async def convert_pdf_to_ppt(file: UploadFile = File(...)):
    # 1. Validasi
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    tmp_pdf_path = None
    tmp_ppt_path = None
    created_images = [] # List untuk melacak gambar sementara agar bisa dihapus

    try:
        # 2. Simpan PDF Sementara
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            shutil.copyfileobj(file.file, tmp_pdf)
            tmp_pdf_path = tmp_pdf.name
        
        tmp_ppt_path = tmp_pdf_path.replace(".pdf", ".pptx")
        original_filename = os.path.splitext(file.filename)[0] + ".pptx"

        # 3. Logic Konversi (PDF -> Image -> Slide)
        prs = Presentation()
        # Set ukuran slide ke Widescreen (16:9) agar modern
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        
        doc = fitz.open(tmp_pdf_path)

        for i, page in enumerate(doc):
            # A. Buat Slide Kosong
            blank_slide_layout = prs.slide_layouts[6] 
            slide = prs.slides.add_slide(blank_slide_layout)
            
            # B. Ubah Halaman PDF jadi Gambar (DPI 150 agar tidak terlalu berat tapi tetap tajam)
            pix = page.get_pixmap(dpi=150)
            img_filename = f"{tmp_pdf_path}_slide_{i}.png"
            pix.save(img_filename)
            created_images.append(img_filename)
            
            # C. Tempel Gambar Full Screen di Slide
            # left=0, top=0, width=full slide width, height=full slide height
            slide.shapes.add_picture(
                img_filename, 
                0, 
                0, 
                width=prs.slide_width, 
                height=prs.slide_height
            )

        # 4. Simpan PPT
        prs.save(tmp_ppt_path)
        doc.close()

        # 5. Kirim File
        return FileResponse(
            path=tmp_ppt_path, 
            filename=original_filename,
            media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        print("!!! ERROR PDF TO PPT !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert PPT: {str(e)}")
        
    finally:
        # Cleanup PDF
        if tmp_pdf_path and os.path.exists(tmp_pdf_path):
            try: os.remove(tmp_pdf_path)
            except: pass
            
        # Cleanup Gambar-gambar sementara (PENTING AGAR SERVER TIDAK PENUH)
        for img_path in created_images:
            if os.path.exists(img_path):
                try: os.remove(img_path)
                except: pass
