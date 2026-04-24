from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import pdfplumber
import pandas as pd
import pytesseract
from PIL import Image
import os

app = FastAPI()

UPLOAD_PATH = "uploads"
OUTPUT_PATH = "outputs"

os.makedirs(UPLOAD_PATH, exist_ok=True)
os.makedirs(OUTPUT_PATH, exist_ok=True)


def extract_with_ocr(pdf_path):
    import fitz  # PyMuPDF
    data = []

    doc = fitz.open(pdf_path)

    for page in doc:
        pix = page.get_pixmap()
        img_path = "temp.png"
        pix.save(img_path)

        text = pytesseract.image_to_string(Image.open(img_path))
        lines = text.split("\n")

        for line in lines:
            data.append([line])

    return pd.DataFrame(data)


@app.post("/convert")
async def convert_pdf(file: UploadFile = File(...)):
    input_path = os.path.join(UPLOAD_PATH, file.filename)
    output_path = os.path.join(OUTPUT_PATH, file.filename.replace(".pdf", ".xlsx"))

    with open(input_path, "wb") as f:
        f.write(await file.read())

    tables = []

    try:
        with pdfplumber.open(input_path) as pdf:
            for page in pdf.pages:
                extracted = page.extract_tables()
                for table in extracted:
                    tables.append(pd.DataFrame(table))
    except:
        pass

    # Se não encontrou tabelas → usa OCR
    if not tables:
        tables.append(extract_with_ocr(input_path))

    with pd.ExcelWriter(output_path) as writer:
        for i, table in enumerate(tables):
            table.to_excel(writer, sheet_name=f"Tabela_{i}")

    return FileResponse(output_path, filename="resultado.xlsx")