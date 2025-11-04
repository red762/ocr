from fastapi import FastAPI, UploadFile, File
from utils.ocr_engine import extract_text_from_file
import tempfile
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="DocQuery OCR API", version="1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # or ["http://localhost:3000"]
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
@app.post("/ocr")
async def ocr_endpoint(file: UploadFile = File(...)):
    try:
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp.write(await file.read())
            tmp.flush()
            extracted = extract_text_from_file(tmp.name, file.filename)
        return {"text": extracted}
    except Exception as e:
        return {"error": str(e)}
