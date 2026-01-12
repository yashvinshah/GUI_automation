# Invoice Processing Automation

Automated invoice processing system that extracts data from PDF invoices and fills Excel and Google Sheets.

##Demo Video

https://github.com/yashvinshah/GUI_automation/releases/download/v1.0.0/video.mp4

## Software Prerequisites

### Required Software
- **Python 3.7 or higher+** - Python programming language
- **Microsoft Excel** (for Excel workflow) - Must be installed and accessible
- **Google Chrome** (for Google Sheets workflow) - Must be installed and accessible

### System Requirements
- **macOS**: Currently fully supported
- **Windows**: Partial support (GUI automation for Windows is in development)

## How to Install Dependencies

1. **Install Python packages** using pip:
   ```bash
   pip install pdfplumber pytesseract pyautogui
   ```

2. **Install Tesseract OCR** (required for OCR fallback):
   
   **macOS:**
   ```bash
   pip install tesseract
   ```
   
   **Windows:**
   - Download from: https://github.com/UB-Mannheim/tesseract/wiki
   - Install the executable and add to PATH


3. **Verify installation:**
   ```bash
   python -c "import pdfplumber, pytesseract, pyautogui; print('All dependencies installed successfully')"
   ```

## How to Prepare Invoice PDFs

1. **Create the invoices directory** (if it doesn't exist):
   ```bash
   mkdir invoices
   ```

2. **Add your invoice PDF files** to the `invoices/` folder:
   - Place all invoice PDF files in the `invoices/` directory
   - Supported formats: `.pdf`, `.PDF`
   - The parser looks for:
     - Invoice Number (pattern: "Invoice Number", "Order No", "Invoice #", "Order #" etc)
     - Invoice Date (pattern: "Invoice Date", "Date of issue", "Order placed" etc)
     - Total Amount (pattern: "Total", "Amount", "Value" followed by numbers)

3. **Configure file paths** (if needed):
   - **Excel file path**: Edit `workflows/excel_workflow.py` and update `EXCEL_FILE_PATH`
   - **Google Sheets URL**: Edit `workflows/google_sheet_workflow.py` and update `GOOGLE_SHEET_URL`

## How to Run the Program

1. **Navigate to the project directory:**
   ```bash
   cd "Your working directory"
   ```

2. **Ensure invoices are in place:**
   - Verify that your invoice PDFs are in the `invoices/` folder

3. **Run the main script:**
   ```bash
   python main.py
   ```

4. **What happens:**
   - The program will parse all PDF invoices in the `invoices/` folder
   - Extracted data is saved to `parsed_invoices.json`
   - Excel workflow opens Excel and fills in invoice data
   - Google Sheets workflow opens Chrome and fills in invoice data
   - Processing status is logged to the console

5. **Output files:**
   - `parsed_invoices.json` - Contains all extracted invoice data
   - `logs/window_manager.log` - Detailed automation logs
   - `logs/` - Screenshots taken during error conditions (stored in logs directory)

## Project Structure

```
GUI_automation/
├── invoice_parser.py          # PDF invoice parsing logic
├── main.py                    # Main entry point
├── invoices/                  # Place your PDF invoices here
├── gui_automation/            # GUI automation code
│   ├── mac_gui.py            # macOS GUI automation
│   ├── windows_gui.py        # Windows GUI automation (placeholder)
│   ├── window_manager.py     # Window management and error handling
│   ├── gui_utils.py          # GUI utility functions
│   └── __init__.py           # Auto-selects platform
└── workflows/                 # Workflow logic
    ├── excel_workflow.py     # Excel automation workflow
    └── google_sheet_workflow.py  # Google Sheets automation workflow
```

## Troubleshooting

- **Excel/Chrome not opening**: Ensure applications are installed and accessible
- **Invoice data not extracted**: Check that invoices contain the expected text patterns
- **Focus issues**: The automation requires the target application to be accessible - avoid moving windows during execution
- **OCR errors**: Ensure Tesseract is properly installed and in your PATH
