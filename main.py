import json
from pathlib import Path
from invoice_parser import parse_invoice
from gui_excel_worker import process_invoice_excel
from gui_googlesheet_worker import process_invoice_google_sheet

# Paths / settings
INVOICES_DIR = Path("invoices")
EXCEL_WORKBOOK_PATH = "/Users/yourusername/Documents/Invoices.xlsx"  # <-- change this
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID"  # <-- change this
LOG_FILE = Path("invoice_processing_log.json")


def main():
    if not INVOICES_DIR.exists():
        print(f"Invoices directory not found: {INVOICES_DIR.resolve()}")
        return

    pdf_files = list(INVOICES_DIR.glob("*.pdf"))

    if not pdf_files:
        print("No PDF files found in invoices directory.")
        return

    print(f"Found {len(pdf_files)} invoice(s)\n")
    results = []

    for pdf_file in pdf_files:
        print(f"Processing: {pdf_file.name}")

        try:
            # Step 1: Parse invoice
            invoice_data = parse_invoice(pdf_file)
            print("Parsed Data:", json.dumps(invoice_data, indent=2))

            # Step 2: Excel GUI automation
            excel_ok = process_invoice_excel(invoice_data, EXCEL_WORKBOOK_PATH)
            print(f"Excel Update: {'Success' if excel_ok else 'Failed'}")

            # Step 3: Google Sheets GUI automation
            google_ok = process_invoice_google_sheet(invoice_data, GOOGLE_SHEET_URL)
            print(f"Google Sheet Update: {'Success' if google_ok else 'Failed'}")

            # Step 4: Log results
            log_entry = {
                "invoice_file": pdf_file.name,
                "invoice_number": invoice_data.get("invoice_number"),
                "excel_status": excel_ok,
                "google_sheet_status": google_ok,
            }
            results.append(log_entry)

            print("-" * 50)

        except Exception as e:
            print(f"Error processing {pdf_file.name}: {e}")

    # Save log
    with open(LOG_FILE, "w") as f:
        json.dump(results, f, indent=2)

    print(f"\nProcessing complete. Log saved to {LOG_FILE.resolve()}")


if __name__ == "__main__":
    main()
