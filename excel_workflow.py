import subprocess
import os
import time
import logging
import re
from pathlib import Path
from typing import List, Dict
import pyautogui
import platform
from gui_utils import wait, press, type_text
from window_manager import WindowManager

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

EXCEL_FILE_PATH = "/Users/nimeshshah/Desktop/NCSU/GUI automation /Invoices.xlsx"
window_mgr = WindowManager("Microsoft Excel", max_retries=3, wait_timeout=10.0)
_excel_opened = False


def extract_amount_numbers(amount_str: str) -> str:
    if not amount_str or amount_str == "null" or str(amount_str).lower() == "none":
        return ""
    
    pattern = r'(\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?|\d+\.\d{1,2}|\d+)'
    matches = re.findall(pattern, str(amount_str))
    
    if matches:
        amount = matches[-1].replace(',', '')
        return amount
    
    pattern_simple = r'(\d+\.?\d*)'
    match = re.search(pattern_simple, str(amount_str))
    if match:
        return match.group(1)
    
    return ""


def clean_date(date_str: str) -> str:
    if not date_str or date_str == "null" or str(date_str).lower() == "none":
        return ""
    
    date_str = str(date_str).strip()
    
    if " " in date_str:
        parts = date_str.split(" ")
        if len(parts) >= 2:
            return parts[0]
    
    return date_str


def open_excel_file() -> bool:
    global _excel_opened
    
    if _excel_opened:
        return window_mgr.focus_application()
    
    if not os.path.exists(EXCEL_FILE_PATH):
        logger.error(f"Excel file not found: {EXCEL_FILE_PATH}")
        return False
    
    logger.info(f"Opening Excel file: {EXCEL_FILE_PATH}")
    
    try:
        subprocess.Popen(["open", "-a", "Microsoft Excel", EXCEL_FILE_PATH])
        
        if window_mgr.wait_for_window_ready(timeout=15.0):
            _excel_opened = True
            wait(2)
            
            if window_mgr.focus_application():
                window_mgr.click_center_safe()
                wait(1)
                logger.info("Excel file opened successfully")
                return True
            else:
                logger.error("Failed to focus Excel window")
                return False
        else:
            logger.error("Excel window did not become ready in time")
            window_mgr.take_screenshot("excel_open_timeout")
            return False
            
    except Exception as e:
        logger.error(f"Error opening Excel: {e}")
        window_mgr.take_screenshot("excel_open_error")
        return False


def ensure_excel_focus() -> bool:
    if not window_mgr.focus_application():
        logger.warning("Failed to focus Excel application")
        return False
    wait(0.3)
    return True


def type_text_to_excel(text: str) -> bool:
    if not ensure_excel_focus():
        logger.error("Excel not focused for typing")
        return False
    
    wait(0.2)
    type_text(str(text))
    return True


def find_invoice_row(invoice_number: str) -> bool:
    logger.info(f"Searching for invoice number: {invoice_number}")
    
    try:
        if not ensure_excel_focus():
            logger.error("Cannot find invoice: Excel not focused")
            return False
        
        wait(0.3)
        
        if platform.system() == "Darwin":
            try:
                script = '''
                tell application "Microsoft Excel"
                    activate
                end tell
                delay 0.2
                tell application "System Events"
                    tell process "Microsoft Excel"
                        keystroke "f" using command down
                    end tell
                end tell
                '''
                subprocess.run(["osascript", "-e", script], check=True, timeout=3)
            except Exception as e:
                logger.warning(f"AppleScript Cmd+F failed: {e}")
                ensure_excel_focus()
                wait(0.2)
                press("command", "f")
        else:
            ensure_excel_focus()
            wait(0.2)
            press("command", "f")
        
        wait(0.2)
        
        type_text(str(invoice_number))
        wait(0.5)
        
        press("enter")
        wait(2.0)
        
        press("escape")
        wait(1.0)
        
        if not ensure_excel_focus():
            logger.error("Lost Excel focus after search")
            return False
        
        logger.info(f"Invoice number {invoice_number} found")
        return True
        
    except Exception as e:
        logger.error(f"Error finding invoice row: {e}")
        window_mgr.take_screenshot(f"find_invoice_error_{invoice_number}")
        try:
            press("escape")
            wait(0.5)
            press("escape")
        except:
            pass
        return False


def update_cells(amount: str, date: str) -> bool:
    logger.info(f"Updating cells - Amount: {amount}, Date: {date}")
    
    try:
        if not ensure_excel_focus():
            logger.error("Excel not focused before updating cells")
            return False
        
        wait(0.5)
        
        press("tab")
        wait(1.5)
        
        if not ensure_excel_focus():
            logger.error("Excel lost focus after Tab")
            return False
        
        type_text_to_excel(str(amount))
        wait(1.0)
        
        if not ensure_excel_focus():
            logger.error("Excel lost focus after typing amount")
            return False
        
        press("tab")
        wait(1.5)
        
        if not ensure_excel_focus():
            logger.error("Excel lost focus after Tab to date")
            return False
        
        type_text_to_excel(str(date))
        wait(1.0)
        
        if not ensure_excel_focus():
            logger.error("Excel lost focus after typing date")
            return False
        
        press("enter")
        wait(0.5)
        
        logger.info("Cells updated successfully")
        return True
        
    except Exception as e:
        logger.error(f"Error updating cells: {e}")
        window_mgr.take_screenshot("update_cells_error")
        return False


def process_all_invoices(invoices_data: List[Dict]) -> Dict[str, bool]:
    results = {}
    
    logger.info(f"Starting to process {len(invoices_data)} invoices")
    
    if not open_excel_file():
        logger.error("Failed to open Excel file")
        for invoice in invoices_data:
            invoice_num = invoice.get("invoice_number", "unknown")
            results[invoice_num] = False
        return results
    
    for idx, invoice_data in enumerate(invoices_data, 1):
        invoice_number = invoice_data.get("invoice_number", f"invoice_{idx}")
        amount = invoice_data.get("total_amount", "")
        date = invoice_data.get("invoice_date", "")
        
        if amount is None:
            amount = ""
        if date is None:
            date = ""
        
        amount = str(amount)
        date = str(date)
        
        if not invoice_number:
            logger.error("Invoice number missing")
            results[invoice_number] = False
            continue
        
        logger.info(f"Processing invoice {idx}/{len(invoices_data)}: {invoice_number}")
        logger.info(f"  Amount: {amount}")
        logger.info(f"  Date: {date}")
        
        if not ensure_excel_focus():
            logger.error("Excel not focused")
            results[invoice_number] = False
            continue
        
        if not find_invoice_row(invoice_number):
            logger.warning(f"Invoice {invoice_number} not found")
            results[invoice_number] = False
            continue
        
        if not update_cells(amount, date):
            logger.error(f"Failed to update cells for invoice {invoice_number}")
            results[invoice_number] = False
            continue
        
        logger.info(f"Successfully processed invoice: {invoice_number}")
        results[invoice_number] = True
        
        if idx < len(invoices_data):
            wait(2.0)
    
    logger.info(f"Completed processing {len(invoices_data)} invoices")
    return results


def process_invoice_excel(invoice_data: Dict) -> bool:
    invoice_number = invoice_data.get("invoice_number")
    amount_raw = invoice_data.get("total_amount", "")
    date_raw = invoice_data.get("invoice_date", "")
    
    amount = extract_amount_numbers(str(amount_raw)) if amount_raw else ""
    date = clean_date(str(date_raw)) if date_raw else ""
    
    if not invoice_number:
        logger.error("Invoice number missing")
        return False
    
    logger.info(f"Processing invoice: {invoice_number}, Amount: {amount}, Date: {date}")
    
    if not ensure_excel_focus():
        logger.error("Excel not focused")
        return False
    
    if not find_invoice_row(invoice_number):
        logger.warning(f"Invoice {invoice_number} not found")
        return False
    
    if not update_cells(amount, date):
        logger.error(f"Failed to update cells for invoice {invoice_number}")
        return False
    
    logger.info(f"Successfully processed invoice: {invoice_number}")
    return True
