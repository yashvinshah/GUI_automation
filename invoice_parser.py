import re
import pdfplumber
import pytesseract
from pathlib import Path

def normalize_text(text: str) -> str:
    text = text.replace("\u00a0", " ")  # non-breaking spaces
    text = re.sub(r"[ \t]+", " ", text) # collapse spaces
    text = re.sub(r"\n+", "\n", text)   # collapse newlines
    return text

# ---------------- Text Extraction ----------------
def extract_text(pdf_path: Path) -> str:
    text = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text.append(page_text)

    combined_text = "\n".join(text)

    if len(combined_text.strip()) < 50:
        ocr_text = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                image = page.to_image(resolution=300).original
                ocr_text.append(pytesseract.image_to_string(image))
        combined_text = "\n".join(ocr_text)

    return combined_text


# ---------------- Invoice Number ----------------
def extract_invoice_number(text: str):
    text = normalize_text(text)

    patterns = [
        r"(I\s*n\s*v\s*o\s*i\s*c\s*e\s*(?:No\.?|Number|#))\s*[:\-]?\s*([A-Za-z0-9\-\/]+)",
        r"(Order\s*(?:No\.?|Number|#|ID))\s*[:\-]?\s*([A-Za-z0-9\-\/]+)",
        r"(Reference\s*(?:No\.?|Number))\s*[:\-]?\s*([A-Za-z0-9\-\/]+)",
        r"(Document\s*(?:No\.?|Number))\s*[:\-]?\s*([A-Za-z0-9\-\/]+)",
    ]

    for p in patterns:
        for m in re.finditer(p, text, re.IGNORECASE):
            value = m.group(2)
            if any(c.isdigit() for c in value):
                return value

    return None

# ---------------- Invoice Date ----------------
def extract_invoice_date(text: str):
    date_patterns = [
        r"(Invoice\s*Date|Date\s*of\s*Issue|Issue\s*Date|Order\s*Placed|Order\s*Date|Date)"
        r"\s*[:\-]?\s*"
        r"("
        r"\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}"
        r"|\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}"
        r"|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}"
        r"|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}"
        r")"
    ]

    for p in date_patterns:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return m.group(2)

    return None


# ---------------- Amount ----------------
def extract_total_amount(text: str):
    amount_pattern = (
        r"(Total|Grand\s*Total|Net\s*Amount|Amount|Balance\s*Due)"
        r"(?:\s*\(.*?\))?"
        r"\s*[:\-]?\s*"
        r"(?:INR|USD|AED|EUR|\$|Rs\.?|₹)?\s*"
        r"([\d,]+(?:\.\d{1,2})?)"
    )

    amounts = []

    for match in re.finditer(amount_pattern, text, re.IGNORECASE):
        amount_str = match.group(2)   # 👈 second capture group
        amounts.append(float(amount_str.replace(",", "")))

    return str(max(amounts)) if amounts else None


# ---------------- Main Parser ----------------
def parse_invoice(pdf_file: str) -> dict:
    text = extract_text(Path(pdf_file))

    return {
        "invoice_file": Path(pdf_file).name,
        "invoice_number": extract_invoice_number(text),
        "invoice_date": extract_invoice_date(text),
        "total_amount": extract_total_amount(text),
    }


if __name__ == "__main__":
    import sys, json

    if len(sys.argv) != 2:
        print("Usage: python invoice_parser.py <invoice.pdf>")
        sys.exit(1)

    result = parse_invoice(sys.argv[1])
    print(json.dumps(result, indent=2))
