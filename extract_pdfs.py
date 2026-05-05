from pathlib import Path
import io
import shutil

import fitz
import pandas as pd
import pytesseract
from pytesseract import TesseractNotFoundError
from PIL import Image


PDF_FOLDER = Path("pdf")
OUTPUT_FOLDER = Path("output")
OUTPUT_FOLDER.mkdir(exist_ok=True)

print("Current folder:", Path.cwd())
print("PDF folder exists:", PDF_FOLDER.exists())
print("PDF files found:", list(PDF_FOLDER.glob("*.pdf")))

# For Windows, uncomment this if Tesseract is installed but not on your PATH:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
OCR_AVAILABLE = shutil.which("tesseract") is not None or Path(
    pytesseract.pytesseract.tesseract_cmd
).exists()

if not OCR_AVAILABLE:
    print("Tesseract OCR was not found. Pages with little/no embedded text will skip OCR.")

rows = []

for pdf_path in PDF_FOLDER.glob("*.pdf"):
    print(f"Processing: {pdf_path.name}")

    try:
        with fitz.open(pdf_path) as doc:
            for page_num, page in enumerate(doc, start=1):
                text = page.get_text("text", sort=True).strip()
                method = "text"

                # If normal extraction finds almost nothing, use OCR.
                if len(text) < 30:
                    if OCR_AVAILABLE:
                        try:
                            pix = page.get_pixmap(dpi=200)
                            img = Image.open(io.BytesIO(pix.tobytes("png")))
                            text = pytesseract.image_to_string(img)
                            method = "ocr"
                        except TesseractNotFoundError as e:
                            OCR_AVAILABLE = False
                            method = "text_ocr_unavailable"
                            text = text or f"OCR skipped: {e}"
                    else:
                        method = "text_ocr_unavailable"
                        text = text or "OCR skipped: Tesseract is not installed or not on PATH."

                rows.append(
                    {
                        "file_name": pdf_path.name,
                        "page_number": page_num,
                        "method": method,
                        "text": text.strip(),
                    }
                )

    except Exception as e:
        rows.append(
            {
                "file_name": pdf_path.name,
                "page_number": None,
                "method": "error",
                "text": f"ERROR: {e}",
            }
        )

df = pd.DataFrame(rows)
df.to_csv(OUTPUT_FOLDER / "extracted_pdf_text.csv", index=False, encoding="utf-8-sig")

print("Done.")
