import subprocess
import os
import time
import logging
import platform
from typing import List, Dict
import pyautogui
from gui_utils import wait, press, type_text
from window_manager import WindowManager

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1K2nESFNfp3zDcM5x-RDCv_pcpdY2ghtcFvillmggVBw/edit?gid=0#gid=0"
window_mgr = WindowManager("Google Chrome", max_retries=3, wait_timeout=10.0)
_sheet_opened = False


def open_google_sheet() -> bool:
    global _sheet_opened
    
    if _sheet_opened:
        return window_mgr.focus_application()
    
    logger.info(f"Opening Google Sheet: {GOOGLE_SHEET_URL}")
    
    try:
        subprocess.Popen(["open", "-a", "Google Chrome", GOOGLE_SHEET_URL])
        
        if window_mgr.wait_for_window_ready(timeout=15.0):
            _sheet_opened = True
            wait(8)
            
            if window_mgr.focus_application():
                window_mgr.click_center_safe()
                wait(1)
                logger.info("Google Sheet opened successfully")
                return True
            else:
                logger.error("Failed to focus Chrome window")
                return False
        else:
            logger.error("Chrome window did not become ready in time")
            window_mgr.take_screenshot("chrome_open_timeout")
            return False
            
    except Exception as e:
        logger.error(f"Error opening Google Sheet: {e}")
        window_mgr.take_screenshot("chrome_open_error")
        return False


def ensure_sheet_focus() -> bool:
    if not window_mgr.focus_application():
        logger.warning("Failed to focus Chrome application")
        return False
    wait(0.3)
    return True


def type_text_to_sheet(text: str) -> bool:
    if not ensure_sheet_focus():
        logger.error("Sheet not focused for typing")
        return False
    
    wait(0.2)
    type_text(str(text))
    return True


def find_invoice_row(invoice_number: str) -> bool:
    logger.info(f"Searching for invoice number: {invoice_number}")
    
    try:
        if not ensure_sheet_focus():
            logger.error("Cannot find invoice: Sheet not focused")
            return False
        
        wait(0.3)
        
        if platform.system() == "Darwin":
            try:
                script = '''
                tell application "Google Chrome"
                    activate
                end tell
                delay 0.2
                tell application "System Events"
                    tell process "Google Chrome"
                        keystroke "f" using command down
                    end tell
                end tell
                '''
                subprocess.run(["osascript", "-e", script], check=True, timeout=3)
            except Exception as e:
                logger.warning(f"AppleScript Cmd+F failed: {e}")
                ensure_sheet_focus()
                wait(0.2)
                press("command", "f")
        else:
            ensure_sheet_focus()
            wait(0.2)
            press("command", "f")
        
        wait(0.2)
        
        type_text(str(invoice_number))
        wait(0.5)
        
        press("enter")
        wait(2.0)
        
        press("escape")
        wait(1.0)
        
        if not ensure_sheet_focus():
            logger.error("Lost sheet focus after search")
            return False
        
        wait(0.3)
        
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
        if not ensure_sheet_focus():
            logger.error("Sheet not focused before updating cells")
            return False
        
        wait(0.5)
        
        press("tab")
        wait(1.5)
        
        if not ensure_sheet_focus():
            logger.error("Sheet lost focus after Tab")
            return False
        
        type_text_to_sheet(str(amount))
        wait(1.0)
        
        if not ensure_sheet_focus():
            logger.error("Sheet lost focus after typing amount")
            return False
        
        press("tab")
        wait(1.5)
        
        if not ensure_sheet_focus():
            logger.error("Sheet lost focus after Tab to date")
            return False
        
        type_text_to_sheet(str(date))
        wait(1.0)
        
        if not ensure_sheet_focus():
            logger.error("Sheet lost focus after typing date")
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
    
    logger.info(f"Starting to process {len(invoices_data)} invoices in Google Sheets")
    
    if not open_google_sheet():
        logger.error("Failed to open Google Sheet")
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
        
        if not ensure_sheet_focus():
            logger.error("Sheet not focused before processing invoice")
            results[invoice_number] = False
            continue
        
        wait(0.3)
        
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
            if not ensure_sheet_focus():
                logger.warning("Sheet lost focus between invoices, refocusing...")
                ensure_sheet_focus()
                wait(0.5)
    
    logger.info(f"Completed processing {len(invoices_data)} invoices")
    return results


def process_invoice_google_sheet(invoice_data: Dict) -> bool:
    invoice_number = invoice_data.get("invoice_number")
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
        return False
    
    logger.info(f"Processing invoice: {invoice_number}, Amount: {amount}, Date: {date}")
    
    if not ensure_sheet_focus():
        logger.error("Sheet not focused")
        return False
    
    if not find_invoice_row(invoice_number):
        logger.warning(f"Invoice {invoice_number} not found")
        return False
    
    if not update_cells(amount, date):
        logger.error(f"Failed to update cells for invoice {invoice_number}")
        return False
    
    logger.info(f"Successfully processed invoice: {invoice_number}")
    return True
