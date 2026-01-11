import sys
import json
from pathlib import Path
from invoice_parser import parse_invoice
from gui_excel_worker import process_all_invoices, process_invoice_excel
from gui_googlesheet_worker import process_all_invoices as process_all_invoices_google, process_invoice_google_sheet

# Directory with invoice PDFs
INVOICE_DIR = "invoices/"

# Log results
results = []

def main():
    invoice_folder = Path(INVOICE_DIR)
    if not invoice_folder.exists():
        print(f"Invoice directory not found: {INVOICE_DIR}")
        sys.exit(1)

    pdf_files = list(invoice_folder.glob("*.pdf"))
    if not pdf_files:
        print("No PDF invoices found.")
        sys.exit(1)

    # Step 1: Parse all invoices first and save to parsed_invoices.json
    all_invoices_data = []
    for pdf_file in pdf_files:
        invoice_data = parse_invoice(str(pdf_file))
        all_invoices_data.append(invoice_data)
    
    # Save parsed invoices to JSON file
    with open("parsed_invoices.json", "w") as f:
        json.dump(all_invoices_data, f, indent=2)

    # Step 2: Process all invoices in a single Excel session
    excel_results = process_all_invoices(all_invoices_data)

    # Step 3: Process all invoices in Google Sheets
    google_results = process_all_invoices_google(all_invoices_data)

    # Step 4: Build results for logging
    for invoice_data in all_invoices_data:
        invoice_number = invoice_data.get("invoice_number", "unknown")
        excel_ok = excel_results.get(invoice_number, False)
        google_ok = google_results.get(invoice_number, False)
        
        results.append({
            "invoice_file": invoice_data.get("invoice_file", "unknown"),
            "invoice_number": invoice_number,
            "excel_status": excel_ok,
            "google_sheet_status": google_ok
        })

if __name__ == "__main__":
    main()
