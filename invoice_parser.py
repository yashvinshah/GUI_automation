import re
import pdfplumber
import pytesseract
from pathlib import Path


def extract_text(pdf_path: Path) -> str:
    text = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text.append(page_text)

    combined_text = "\n".join(text)

    # Fallback to OCR if text extraction fails
    if len(combined_text.strip()) < 50:
        ocr_text = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                image = page.to_image(resolution=300).original
                ocr_text.append(pytesseract.image_to_string(image))
        combined_text = "\n".join(ocr_text)

    return combined_text


def parse_invoice(pdf_file: str) -> dict:
    pdf_path = Path(pdf_file)
    text = extract_text(pdf_path)

    data = {
        "invoice_file": pdf_path.name,
        "invoice_number": None,
        "invoice_date": None,
        "total_amount": None,
    }

    # Simple regex rules (edit per invoice format)
    rules = {
    "invoice_number": r"(Invoice\s*(Number|No\.?|#)\s*[:\-]?\s*)([^\n\r]+)",

    "invoice_date": r"(Invoice\s*Date\s*[:\-]?\s*)([^\n\r]+)",

    "total_amount": r"(Total|Amount\s*Due)\s*[:$]?\s*([^\n\r]+)",
}

    for field, pattern in rules.items():
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data[field] = match.groups()[-1]

    return data


if __name__ == "__main__":
    import sys
    import json

    if len(sys.argv) != 2:
        print("Usage: python invoice_parser.py <invoice.pdf>")
        sys.exit(1)

    result = parse_invoice(sys.argv[1])
    print(json.dumps(result, indent=2))
