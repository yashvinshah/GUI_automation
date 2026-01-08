import subprocess
from gui_utils import wait, press, type_text
import pyautogui


def open_excel():
    subprocess.Popen(["open", "-a", "Microsoft Excel"])
    wait(5)


def open_workbook(file_path):
    press("command", "o")
    wait(1)
    type_text(file_path)
    press("enter")
    wait(3)


def find_invoice_row(invoice_number):
    press("command", "f")
    wait(1)
    type_text(invoice_number)
    press("enter")
    wait(1)
    press("escape")  # close find box
    wait(0.5)


def update_cells(amount, date):
    # Assumes cursor is on invoice row
    press("right")      # move to amount column
    type_text(amount)

    press("right")      # move to date column
    type_text(date)

    press("enter")


def process_invoice_excel(invoice_data, workbook_path):
    try:
        open_excel()
        open_workbook(workbook_path)

        find_invoice_row(invoice_data["invoice_number"])
        update_cells(
            invoice_data["total_amount"],
            invoice_data["invoice_date"],
        )

        return True

    except Exception as e:
        print(f"[EXCEL ERROR] {e}")
        return False
