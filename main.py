import os
import shutil
import tempfile
import traceback
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

@app.post("/convert/pdf-to-docx")
async def convert_pdf(file: UploadFile = File(...)):
    # 1. Validasi
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File harus format PDF")

    # Variabel path sementara
    tmp_pdf_path = None
    tmp_docx_path = None

    try:
        # 2. Buat File Sementara (Temporary)
        # Ini lebih aman daripada membuat folder manual di server
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            shutil.copyfileobj(file.file, tmp_pdf)
            tmp_pdf_path = tmp_pdf.name
        
        # Tentukan nama output docx di folder temp yang sama
        tmp_docx_path = tmp_pdf_path.replace(".pdf", ".docx")
        
        # Nama file asli untuk didownload user
        original_filename = os.path.splitext(file.filename)[0] + ".docx"

        # 3. PROSES KONVERSI (SAFE MODE)
        # multiprocess=False -> Agar tidak memberatkan CPU Render
        # cpu_count=1 -> Paksa pakai 1 core saja
        cv = Converter(tmp_pdf_path)
        cv.convert(tmp_docx_path, start=0, end=None, multiprocess=False, cpu_count=1)
        cv.close()

        # 4. Kirim File
        return FileResponse(
            path=tmp_docx_path, 
            filename=original_filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        # Cetak error lengkap ke Log Render untuk debugging
        print("!!! ERROR TERJADI SAAT KONVERSI !!!")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Gagal memproses file: {str(e)}")
        
    finally:
        # 5. Cleanup (Hapus file sampah PDF input)
        # Kita biarkan file DOCX dihapus otomatis oleh sistem nanti atau overwrite
        if tmp_pdf_path and os.path.exists(tmp_pdf_path):
            try:
                os.remove(tmp_pdf_path)
            except:
                pass
