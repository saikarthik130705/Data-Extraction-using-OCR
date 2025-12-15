import os
import re
import cv2
import pytesseract
import pdfplumber
import pandas as pd
from pdf2image import convert_from_path

# -------- CONFIG --------
INVOICE_FOLDER = "invoices\Chinese"
OUTPUT_EXCEL = "extracted_invoices_chinese.xlsx"

# English + Chinese + Danish
TESSERACT_LANGS = "eng+chi_sim+dan"

# -------- IMAGE PREPROCESSING --------
def preprocess_image(image_path):
    img = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    img = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY)[1]
    return img

# -------- EXTRACT TEXT FROM PDF --------
def extract_text_from_pdf(pdf_path):
    text = ""

    # Try text-based extraction
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"

    # OCR fallback
    if len(text.strip()) < 20:
        pages = convert_from_path(pdf_path)
        for i, page in enumerate(pages):
            img_path = f"temp_{i}.png"
            page.save(img_path)

            processed = preprocess_image(img_path)
            ocr_text = pytesseract.image_to_string(
                processed,
                lang=TESSERACT_LANGS
            )
            text += ocr_text + "\n"
            os.remove(img_path)

    return text

# -------- FIELD EXTRACTION --------
def extract_invoice_fields(text):
    return {
        "Invoice Number": re.search(r"(Invoice|发票|Faktura).*?(\w+)", text, re.I).group(2)
        if re.search(r"(Invoice|发票|Faktura).*?(\w+)", text, re.I) else None,

        "Date": re.search(r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})", text).group(1)
        if re.search(r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})", text) else None,

        "Total Amount": re.search(r"([\d,.]+\.\d{2})", text).group(1)
        if re.search(r"([\d,.]+\.\d{2})", text) else None
    }

# -------- MAIN --------
all_data = []

for file in os.listdir(INVOICE_FOLDER):
    if file.endswith(".pdf"):
        print(f"Processing: {file}")
        pdf_path = os.path.join(INVOICE_FOLDER, file)

        text = extract_text_from_pdf(pdf_path)
        fields = extract_invoice_fields(text)
        fields["Source File"] = file

        all_data.append(fields)

df = pd.DataFrame(all_data)
df.to_excel(OUTPUT_EXCEL, index=False)

print("✅ Extraction completed. Excel created.")