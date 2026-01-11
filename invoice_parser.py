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

    # Invoice number
    print("Searching for invoice number...")
    invoice_number_pattern = r"(Invoice\s*(Number|No\.?|#)\s*[:\-]?\s*)([^\n\r]+)"
    match = re.search(invoice_number_pattern, text, re.IGNORECASE)
    if match:
        data["invoice_number"] = match.groups()[-1].strip()
        print(f"  ✓ Found invoice number: {data['invoice_number']}")
    else:
        print("  ✗ Invoice number not found")

    # Invoice date
    print("Searching for invoice date...")
    invoice_date_pattern = r"(Invoice\s*Date\s*[:\-]?\s*)([^\n\r]+)"
    match = re.search(invoice_date_pattern, text, re.IGNORECASE)
    if match:
        data["invoice_date"] = match.groups()[-1].strip()
        print(f"  ✓ Found invoice date: {data['invoice_date']}")
    else:
        print("  ✗ Invoice date not found")

    # Amount - look for "Total" or "Amount" followed by numbers
    print("Searching for total amount...")
    amount_pattern = r"(Total|Amount)\s*[:\-]?\s*(?:INR|USD|\$|Rs\.?|₹)?\s*([\d,]+(?:\.\d{1,2})?)"
    match = re.search(amount_pattern, text, re.IGNORECASE)
    if match:
        data["total_amount"] = match.group(2).replace(',', '').strip()
        print(f"  ✓ Found total amount: {data['total_amount']}")
    else:
        print("  ✗ Total amount not found")

    return data


if __name__ == "__main__":
    import sys
    import json

    if len(sys.argv) != 2:
        print("Usage: python invoice_parser.py <invoice.pdf>")
        sys.exit(1)

    result = parse_invoice(sys.argv[1])
    print(json.dumps(result, indent=2))
