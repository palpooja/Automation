"""engines/sheet_safety.py - keep text from the portal as text in Excel and CSV files.

Excel (and openpyxl) treat a cell value that starts with '=' as a formula.  Text
copied from the portal - for example a withholding agent's name typed by another
taxpayer - must never become a live formula in a client's workbook (finding L2,
the risk OWASP calls 'CSV injection').  The program's own formulas (the TOTAL
=SUM rows) are written separately and are not touched.
"""
import re

_TRIGGERS = ("=", "+", "-", "@", "\t", "\r")
_NUMBER = re.compile(r"[+-]?[\d,]*\.?\d+")


def is_risky(value) -> bool:
    """True for text that a spreadsheet could run as a formula."""
    if not isinstance(value, str) or not value or value[0] not in _TRIGGERS:
        return False
    return not _NUMBER.fullmatch(value.strip())      # '-500,000' is a number, not a formula


def as_text_cell(cell):
    """After writing a value into an openpyxl cell: if it is risky text, store it as text."""
    if is_risky(cell.value):
        cell.data_type = "s"
    return cell


def neutralise_sheet(ws):
    """For sheets written by pandas.to_excel (no formulas of our own): every risky text stays text."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type == "f" or is_risky(cell.value):
                cell.data_type = "s"


def csv_text(value):
    """For CSV files: prefix a quote, as OWASP recommends, so Excel shows the text."""
    return "'" + value if is_risky(value) else value
