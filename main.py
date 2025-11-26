import os
import shutil
import tempfile
import traceback
import pdfplumber
import pandas as pd
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter

app = FastAPI()

# --- CORS ---
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- HOME ---
@app.get("/")
def home():
    return {"status": "Server Online", "endpoints": ["/convert/pdf-to-docx", "/convert/pdf-to-excel"]}

# === FITUR 1: PDF KE WORD ===
@app.post("/convert/pdf-to-docx")
async def convert_pdf_to_docx(file: UploadFile = File(...)):
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

        return FileResponse(
            path=tmp_docx_path, 
            filename=original_filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        print("!!! ERROR PDF TO DOCX !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Word: {str(e)}")
        
    finally:
        if tmp_pdf_path and os.path.exists(tmp_pdf_path):
            try: os.remove(tmp_pdf_path)
            except: pass

# === FITUR 2: PDF KE EXCEL (BARU) ===
@app.post("/convert/pdf-to-excel")
async def convert_pdf_to_excel(file: UploadFile = File(...)):
    # 1. Validasi
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    tmp_pdf_path = None
    tmp_xlsx_path = None

    try:
        # 2. Simpan PDF Sementara
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            shutil.copyfileobj(file.file, tmp_pdf)
            tmp_pdf_path = tmp_pdf.name
        
        tmp_xlsx_path = tmp_pdf_path.replace(".pdf", ".xlsx")
        original_filename = os.path.splitext(file.filename)[0] + ".xlsx"

        # 3. PROSES KONVERSI (Logic Ekstraksi Tabel)
        with pdfplumber.open(tmp_pdf_path) as pdf:
            # Siapkan writer untuk menulis file Excel
            with pd.ExcelWriter(tmp_xlsx_path, engine='openpyxl') as writer:
                tables_found = False
                
                # Loop setiap halaman PDF
                for i, page in enumerate(pdf.pages):
                    # Cari tabel di halaman tersebut
                    tables = page.extract_tables()
                    
                    if tables:
                        tables_found = True
                        row_tracker = 0 # Agar tabel tidak menumpuk
                        
                        # Loop setiap tabel yang ketemu di satu halaman
                        for table in tables:
                            df = pd.DataFrame(table)
                            # Simpan ke Sheet bernama "Page X"
                            sheet_name = f"Page {i+1}"
                            
                            # Tulis ke Excel
                            df.to_excel(writer, sheet_name=sheet_name, startrow=row_tracker, index=False, header=False)
                            
                            # Beri jarak 2 baris untuk tabel berikutnya (jika ada)
                            row_tracker += len(df) + 2

                # Jika PDF isinya teks semua (bukan tabel), buat sheet info
                if not tables_found:
                    df_fallback = pd.DataFrame(["Maaf, tidak ditemukan struktur tabel yang jelas di PDF ini."])
                    df_fallback.to_excel(writer, sheet_name="Info", index=False, header=False)

        # 4. Kirim File Excel
        return FileResponse(
            path=tmp_xlsx_path, 
            filename=original_filename,
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        print("!!! ERROR PDF TO EXCEL !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal convert Excel: {str(e)}")
        
    finally:
        # Cleanup PDF input (Excel output biarkan dihapus sistem nanti)
        if tmp_pdf_path and os.path.exists(tmp_pdf_path):
            try: os.remove(tmp_pdf_path)
            except: pass
