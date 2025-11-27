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

# === FITUR 4: PDF KE GAMBAR (JPG, JPEG, PNG Only) ===
@app.post("/convert/pdf-to-image")
async def convert_pdf_to_image(
    output_format: str = "png", # Default PNG
    file: UploadFile = File(...)
):
    # 1. Validasi Format (GIF DIHAPUS)
    fmt = output_format.lower() 
    # Hanya izinkan 3 format ini
    allowed_formats = ["jpg", "jpeg", "png"]
    
    if fmt not in allowed_formats:
        # Error 400: Bad Request (Salah Format)
        raise HTTPException(status_code=400, detail="Format tidak didukung. Pilih: jpg, jpeg, atau png.")
        
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    # Setup Folder Sementara
    tmp_dir = tempfile.mkdtemp() 
    tmp_pdf_path = os.path.join(tmp_dir, file.filename)
    
    # Nama ZIP output
    zip_filename = os.path.splitext(file.filename)[0] + f"_images_{fmt}.zip"
    tmp_zip_path = os.path.join(tmp_dir, zip_filename)
    
    images_folder = os.path.join(tmp_dir, "processed_images")
    os.makedirs(images_folder, exist_ok=True)

    try:
        # 2. Simpan PDF
        with open(tmp_pdf_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = fitz.open(tmp_pdf_path)

        # Logic Alpha Channel (Transparansi)
        # JPG/JPEG tidak support transparan, jadi harus False
        use_alpha = False if fmt in ["jpg", "jpeg"] else True

        # 3. Loop Convert
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=150, alpha=use_alpha)
            
            page_number = str(i + 1).zfill(3) 
            img_name = f"page_{page_number}.{fmt}"
            img_path = os.path.join(images_folder, img_name)
            
            pix.save(img_path)

        doc.close()

        # 4. ZIP Semua Gambar
        with zipfile.ZipFile(tmp_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(images_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, arcname=file)

        # 5. Kirim File ZIP
        return FileResponse(
            path=tmp_zip_path, 
            filename=zip_filename,
            media_type='application/zip'
        )

    except Exception as e:
        print("!!! ERROR PDF TO IMAGE !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Image: {str(e)}")
        
    finally:
        try: shutil.rmtree(tmp_dir)
        except: pass
