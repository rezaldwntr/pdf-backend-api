import os
import shutil
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter
import traceback # Tambahkan import ini di paling atas

app = FastAPI()

# --- SETUP CORS (Wajib agar Frontend bisa akses) ---
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Mengizinkan semua domain (termasuk Vercel)
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

def cleanup_files(paths):
    for path in paths:
        if os.path.exists(path):
            os.remove(path)

@app.post("/convert/pdf-to-docx")
async def convert_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="File must be PDF")

    input_path = os.path.join(UPLOAD_DIR, file.filename)
    output_filename = os.path.splitext(file.filename)[0] + ".docx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Proses Konversi
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()

        # Hapus file setelah response dikirim
        background_tasks.add_task(cleanup_files, [input_path, output_path])

        return FileResponse(
            path=output_path, 
            filename=output_filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        cleanup_files([input_path])
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
def root():
    return {"message": "PDF Converter API is Running on Render!"}

    try:
        # Simpan file PDF
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Proses Konversi
        # UPDATE DISINI: Tambahkan settings custom untuk mematikan multiprocessing
        cv = Converter(input_path)
        
        # cpu_count=1 dan multiprocess=False mencegah Render crash
        cv.convert(output_path, start=0, end=None, multiprocess=False, cpu_count=1) 
        cv.close()

        background_tasks.add_task(cleanup_files, [input_path, output_path])

        return FileResponse(
            path=output_path, 
            filename=output_filename,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        # Tampilkan error lengkap di Log Render (Traceback)
        print("!!! ERROR DETAIL !!!")
        traceback.print_exc() 
        
        cleanup_files([input_path])
        raise HTTPException(status_code=500, detail=f"Server Error: {str(e)}")
