"""
URA (Uganda Revenue Authority) e-Tax Returns Downloader  -  version 3
======================================================================
WHAT CHANGED IN VERSION 4.3
1. Excel status sheet: every client folder now also gets
   Returns_Status.xlsx. Its 'Returns' tab holds one line per return the
   portal listed, for every ticked return type, copied exactly from the
   Return History table (including its status and original/amended
   columns, under the portal's own headings) plus 'Return type',
   'Downloaded' and 'Checked on'. Its 'Summary' tab counts the returns
   of each type per status, so amended returns still pending approval
   are visible at a glance. Later runs merge into the same file: a
   return already listed is updated (a status that changed is
   refreshed), new ones are added, old ones kept. File names of the
   PDFs are unchanged. Needs the 'openpyxl' package (the build scripts
   install and bundle it); without it the sheet is skipped with a note.

WHAT CHANGED IN VERSION 4.2
1. Headquarter / Both: for Local Excise Duty, Monthly VAT, PAYE and
   Withholding Tax the portal shows a second drop-down, preset to
   'Headquarter', after the return type is chosen. The bot now switches
   it to 'Both' before pressing Search, so returns filed under every
   branch are listed and downloaded too. The log shows the switch. The
   box is looked for ONLY on those four types (BRANCH_CODES), so the
   other thirteen types run exactly as before with no extra wait.

WHAT CHANGED IN VERSION 4.1
1. Print button: if a return has 'Print with Schedules' (VAT, PAYE,
   Withholding Tax ...) that button is ALWAYS used, so the schedules are
   in the PDF. Plain 'Print' is used only when the return has no
   'Print with Schedules' button. (Before, whichever button came first
   on the page was pressed - and plain 'Print' comes first.) The log
   now says which button was used for every return.

WHAT CHANGED IN VERSION 4
1. Saved clients: keep any number of Login ID + password pairs (passwords
   are encrypted with Windows' own user-level protection, DPAPI), tick the
   ones you want, press Fetch once. Chrome opens for each client in turn;
   you type the CAPTCHA each time.
2. History: every run is recorded (when, client, period, files saved) and
   can be viewed from the History button.

WHAT CHANGED IN VERSION 3
1. "Wait until it downloads" tick box. Slow returns used to be skipped
   after short fixed waits (12 s for the tab, 14 s for the Print button,
   40 s for the PDF counted from the first click). Now the PDF clock only
   starts AFTER Print is pressed, the normal waits are longer, and with
   the tick box ON the bot keeps waiting for each return (up to a
   15-minute safety limit) and retries failed rows three times.
2. Local Excise Duty Return: no PDF - instead a faithful copy of the
   web page is saved as .mhtml (a single-file web archive that keeps the
   text, tables, styling and every embedded frame exactly as shown on
   screen). Double-click it to open in Chrome/Edge.

WHAT CHANGED IN VERSION 2
The URA portal hides its content inside frames-within-frames (embedded
pages inside embedded pages). Version 1 only looked inside the first
frame and could miss things. Version 2 searches EVERY frame at EVERY
level until it finds what it needs, and logs each step in detail.

HOW IT WORKS
1. Fill in Login ID, Password, From date, To date -> press FETCH
2. Chrome opens, bot fills your login details
3. YOU type the security code (CAPTCHA) and click "Log In"
4. Bot skips popups, opens "View return history", fills dates,
   loops every return type, opens every return, silently saves each
   as PDF and files it into  Downloads/URA_Returns/<Category>/
"""

import os
import re
import sys
import queue
import json
import time
import shutil
import threading
import base64
import ctypes
import datetime
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.common.exceptions import (TimeoutException,
                                        NoSuchWindowException,
                                        WebDriverException)

try:                              # for the Returns_Status.xlsx sheet
    import openpyxl
    from openpyxl.styles import Font
except Exception:                 # the sheet is optional - never stop the bot
    openpyxl = None

BASE_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads", "URA_Returns")
# Saved clients and run history live in the user's private AppData folder
DATA_FOLDER = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")),
                           "URA_Returns_Downloader")
CLIENTS_FILE = os.path.join(DATA_FOLDER, "clients.json")
HISTORY_FILE = os.path.join(DATA_FOLDER, "history.json")
TEMP_PRINT_FOLDER = os.path.join(BASE_FOLDER, "_temp_prints")
LOGIN_URL = "https://ura.go.ug/en/etax-login/"
PDF_WAIT_SECONDS = 90              # normal mode: after Print is clicked
TAB_WAIT_SECONDS = 60              # normal mode: for the return's tab to open
PAGE_WAIT_SECONDS = 60             # normal mode: for the Print button to appear
PATIENT_LIMIT_SECONDS = 15 * 60    # "wait until it downloads" safety valve
HTML_CATEGORIES = {"Local Excise Duty Return"}   # also saved as .mhtml
STATUS_XLSX = "Returns_Status.xlsx"            # one per client, all return types

# Only these four return types show the Headquarter/Both box on the
# return-history page (portal codes, as seen in the user's recording).
# The bot looks for the box ONLY for these, so no time is wasted on the
# other types.
BRANCH_CODES = {
    "LEDRT",     # Local Excise Duty Return
    "VATRT",     # Monthly VAT Return
    "PAYERT",    # PAYE Return
    "WTHDRT",    # Withholding Tax Return
}

# The 17 return types the URA portal offers (taken from a real run).
# These become tick boxes in the app window.
ALL_CATEGORIES = [
    "Diplomatic VAT Return",
    "Gaming Tax Tax Return",
    "Gaming Tax Tax Return - Weekly",
    "Income Tax Provisional Return for Individual",
    "Income Tax Return for Individual",
    "Income Tax Return for Individual with Employment/Rental Income",
    "Income Tax Return for Non-Individual",
    "Income Tax Return for Partnership",
    "Income Tax Return for Presumptive Taxpayer",
    "Local Excise Duty Return",
    "Monthly VAT Return",
    "NON-RESIDENT QUARTERLY TAX RETURN",
    "PAYE Return",
    "Provisional Return for Non-Individual",
    "Rental Income Tax Return",
    "Return filing from the Bulk Assessment Authority Agent.",
    "Withholding Tax Return",
]

# These come TICKED by default (the user's usual needs); others start unticked
DEFAULT_TICKED = {
    "Monthly VAT Return",
    "PAYE Return",
    "Income Tax Return for Individual",
    "Income Tax Return for Non-Individual",
    "Income Tax Provisional Return for Individual",
    "Provisional Return for Non-Individual",
    "Local Excise Duty Return",
}


# ======================================================================
# SAVED CLIENTS + HISTORY (small JSON files, passwords DPAPI-encrypted)
# ======================================================================
def _dpapi(data: bytes, protect: bool) -> bytes:
    """Encrypt/decrypt with Windows DPAPI (Data Protection API) - the
    secret can only be read back by the same Windows user on the same PC.
    Outside Windows it falls back to plain bytes (dev use only)."""
    if os.name != "nt":
        return data

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", ctypes.c_uint32),
                    ("pbData", ctypes.POINTER(ctypes.c_char))]
    crypt = ctypes.windll.crypt32
    kernel = ctypes.windll.kernel32
    buf = ctypes.create_string_buffer(data, len(data))
    inp = BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    out = BLOB()
    fn = crypt.CryptProtectData if protect else crypt.CryptUnprotectData
    if not fn(ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out)):
        raise OSError("DPAPI call failed")
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        kernel.LocalFree(out.pbData)


def protect_secret(text: str) -> str:
    return base64.b64encode(_dpapi(text.encode("utf-8"), True)).decode("ascii")


def unprotect_secret(blob: str) -> str:
    try:
        return _dpapi(base64.b64decode(blob), False).decode("utf-8")
    except Exception:
        return ""


def _load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return default


def _save_json(path, data):
    os.makedirs(DATA_FOLDER, exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2)
    os.replace(tmp, path)


def load_clients():
    """[{name, login, password}] - password decrypted for use."""
    out = []
    for c in _load_json(CLIENTS_FILE, []):
        out.append({"name": c.get("name", ""), "login": c.get("login", ""),
                    "password": unprotect_secret(c.get("password", ""))})
    return out


def save_clients(clients):
    _save_json(CLIENTS_FILE, [
        {"name": c["name"], "login": c["login"],
         "password": protect_secret(c["password"])} for c in clients])


def load_history():
    return _load_json(HISTORY_FILE, [])


def add_history(entry):
    hist = load_history()
    hist.append(entry)
    _save_json(HISTORY_FILE, hist[-500:])       # keep the last 500 runs


# ======================================================================
# BROWSER SETUP (silent "Save as PDF" printing)
# ======================================================================
def make_driver():
    os.makedirs(TEMP_PRINT_FOLDER, exist_ok=True)
    options = webdriver.ChromeOptions()
    print_settings = {
        "recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
        "selectedDestinationId": "Save as PDF",
        "version": 2,
    }
    options.add_experimental_option("prefs", {
        "printing.print_preview_sticky_settings.appState": json.dumps(print_settings),
        "savefile.default_directory": TEMP_PRINT_FOLDER,
        "download.default_directory": TEMP_PRINT_FOLDER,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    })
    # hide the "Chrome is being controlled by automated test software" bar
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--kiosk-printing")
    options.add_argument("--start-maximized")
    return webdriver.Chrome(options=options)


# ======================================================================
# FRAME-SEARCHING HELPERS  (the important new part)
# ======================================================================
# Memory of which frame each element was found in, so repeat lookups
# jump straight there instead of sweeping every frame again
_frame_memory = {}


def _try_remembered_frame(driver, by, sel, require_visible):
    """If we found this element before, go straight to its frame."""
    path = _frame_memory.get(sel)
    if path is None:
        return None
    try:
        driver.switch_to.default_content()
        for idx in path:
            frames = driver.find_elements(By.XPATH, "//iframe | //frame")
            if idx >= len(frames):
                return None
            driver.switch_to.frame(frames[idx])
        for el in driver.find_elements(by, sel):
            try:
                if el.is_displayed() or not require_visible:
                    return el
            except Exception:
                continue
    except Exception:
        pass
    return None


def _search_all_frames(driver, by, sel, depth=0, max_depth=3,
                       require_visible=True, path=None):
    """Look for the element here, then inside every frame, then inside
    every frame's frames... Returns the element, LEAVES the browser
    focused on the frame containing it, and REMEMBERS that frame.
    require_visible=False also accepts elements hidden inside folded
    menus (needed for links that only appear when you hover a menu)."""
    if path is None:
        path = []
    try:
        for el in driver.find_elements(by, sel):
            try:
                if el.is_displayed() or not require_visible:
                    _frame_memory[sel] = list(path)
                    return el
            except Exception:
                continue
    except Exception:
        pass
    if depth >= max_depth:
        return None
    try:
        frame_count = len(driver.find_elements(By.XPATH, "//iframe | //frame"))
    except Exception:
        return None
    for i in range(frame_count):
        try:
            frames = driver.find_elements(By.XPATH, "//iframe | //frame")
            if i >= len(frames):
                break
            driver.switch_to.frame(frames[i])
        except Exception:
            continue
        path.append(i)
        found = _search_all_frames(driver, by, sel, depth + 1, max_depth,
                                   require_visible, path)
        if found is not None:
            return found
        path.pop()
        try:
            driver.switch_to.parent_frame()
        except Exception:
            driver.switch_to.default_content()
    return None


def locate(driver, by, sel, timeout=10, require_visible=True):
    """Find an element: first try the remembered frame (fast), then do a
    full sweep of all frames (slow), retrying up to `timeout` seconds."""
    end = time.time() + timeout
    while True:
        el = _try_remembered_frame(driver, by, sel, require_visible)
        if el is not None:
            return el
        driver.switch_to.default_content()
        el = _search_all_frames(driver, by, sel, require_visible=require_visible)
        if el is not None:
            return el
        if time.time() > end:
            return None
        time.sleep(0.25)


def js_click(driver, element):
    driver.execute_script("arguments[0].click();", element)


def newest_pdf(folder, since_time):
    try:
        for name in os.listdir(folder):
            if name.lower().endswith(".pdf"):
                full = os.path.join(folder, name)
                if os.path.getmtime(full) >= since_time:
                    size1 = os.path.getsize(full)
                    time.sleep(0.5)
                    if os.path.getsize(full) == size1:
                        return full
    except FileNotFoundError:
        pass
    return None


def safe_name(text):
    return "".join(c if c.isalnum() or c in " -_()." else "_" for c in text).strip()


def unique_path(folder, stem, ext):
    """<stem>.ext, or <stem>_2.ext, _3... if that name is already taken."""
    target = os.path.join(folder, f"{stem}{ext}")
    n = 2
    while os.path.exists(target):
        target = os.path.join(folder, f"{stem}_{n}{ext}")
        n += 1
    return target


def fill_dates(driver, date_from, date_to, log):
    """Fill the two date boxes. The 'From' box's confirmed name is
    #fromDate (seen in the user's recording); its partner is tried as
    #toDate. If either is missing, fall back to hunting date-style
    text boxes, then to the first two text boxes on the form."""
    anchor = locate(driver, By.ID, "rtnType", 15)   # positions us in the right frame
    if anchor is None:
        return False

    boxes = []
    named_from = driver.find_elements(By.ID, "fromDate")
    named_to = driver.find_elements(By.ID, "toDate")
    if named_from and named_to:
        boxes = [named_from[0], named_to[0]]
    if len(boxes) < 2:
        boxes = driver.find_elements(
            By.XPATH,
            "//input[@type='text' and ("
            "contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'date') or "
            "contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'date'))]")
    if len(boxes) < 2:
        boxes = [b for b in driver.find_elements(By.XPATH, "//input[@type='text']")
                 if b.is_displayed()]
    if len(boxes) < 2:
        log("   Could not find the date boxes in any frame.")
        return False
    for box, value in ((boxes[0], date_from), (boxes[1], date_to)):
        driver.execute_script(
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
            box, value)
    return True


def select_both_branches(driver, log, timeout=4):
    """For Excise, VAT, PAYE and Withholding Tax (see BRANCH_CODES) the
    portal shows a second drop-down #branchDrpDwn, preset to
    'Headquarter', once the return type is chosen. Switch it to 'Both'
    (option value 'B', as seen in the user's recording) so returns filed
    under every branch are listed as well. Called only for those four
    types; the box normally appears at once, so the timeout is just a
    safety ceiling.

    Returns True when 'Both' is selected, False when the box is not on
    this screen (or could not be switched)."""
    dd = locate(driver, By.ID, "branchDrpDwn", timeout)   # visible ones only
    if dd is None:
        log("   Headquarter/Both box did not appear - continuing with "
            "the portal's default (Headquarter).")
        return False
    sel = Select(dd)
    try:
        before = sel.first_selected_option.text.strip() or "(blank)"
    except Exception:
        before = "?"
    if before.lower() == "both":
        log("   Headquarter/Both box is already 'Both'.")
        return True
    try:
        sel.select_by_value("B")
    except Exception:
        # in case URA ever changes the code letter: pick it by the word
        try:
            for i, opt in enumerate(sel.options):
                if opt.text.strip().lower() == "both":
                    sel.select_by_index(i)
                    break
        except Exception:
            pass
    try:
        after = sel.first_selected_option.text.strip()
    except Exception:
        after = "?"
    if after.lower() == "both":
        log(f"   Headquarter/Both box: '{before}' -> 'Both'")
        return True
    log(f"   Could not switch the Headquarter/Both box to 'Both' "
        f"(it still says '{after}') - continuing with '{after}'.")
    return False



# ======================================================================
# EXCEL STATUS SHEET  (Returns_Status.xlsx - one per client)
# ======================================================================
# Every row of the portal's Return History table is copied, exactly as
# shown, into one Excel file per client, together with the return type
# the bot searched. Whatever URA calls the status column ('Status',
# 'Return Status' ...) it lands in the sheet under that heading, so the
# user can see at a glance which amended returns are approved and which
# are still pending. A Summary tab counts them per return type.
BOT_TYPE, BOT_DL, BOT_CHK = "Return type", "Downloaded", "Checked on"


def _clean(text):
    return " ".join((text or "").split())


def read_history_headings(driver):
    """The ten column headings of the Return History table, read from
    its header row (the 11th column - the View link - is left out).
    Falls back to 'Column 1' .. 'Column 10'. The caller has already
    positioned the driver in the table's frame with locate()."""
    heads = []
    for xp in ("(//*[@id='returnHistory']//tr[th])[1]/th",
               "(//*[@id='returnHistory']//thead//tr)[1]/td"):
        try:
            cells = driver.find_elements(By.XPATH, xp)
        except Exception:
            cells = []
        if cells:
            for c in cells:
                t = _clean(c.text)
                if not t:
                    t = _clean(c.get_attribute("textContent"))
                heads.append(t)
            break
    heads = heads[:10] + [""] * max(0, 10 - len(heads))
    out, seen = [], set()
    for i, h in enumerate(heads):
        h = h or f"Column {i + 1}"
        base, n = h, 2
        while h in seen or h in (BOT_TYPE, BOT_DL, BOT_CHK):
            h = f"{base} ({n})"
            n += 1
        seen.add(h)
        out.append(h)
    return out


def read_row_cells(row):
    """Text of the first ten cells of one Return History row."""
    tds = row.find_elements(By.XPATH, "./td")
    if not tds:
        tds = row.find_elements(By.XPATH, ".//td")
    out = []
    for td in tds[:10]:
        t = _clean(td.text)
        if not t:
            try:
                t = _clean(td.get_attribute("textContent"))
            except Exception:
                t = ""
        out.append(t)
    return out + [""] * max(0, 10 - len(out))


def write_status_sheet(tin_folder, entries, checked_on, log):
    """Write (or merge into) <tin_folder>/Returns_Status.xlsx.

    entries: [{"type", "heads", "cells", "downloaded"}] from this run.
    An existing file is loaded and merged - same return (type + return
    number) is replaced so a changed status is refreshed; other returns
    from earlier runs are kept. The Summary tab is rebuilt from the full
    list. Returns a dict {path, status_head, breakdown} or None."""
    if not entries:
        return None
    if openpyxl is None:
        log("   Status sheet skipped - the 'openpyxl' package is not "
            "installed (pip install openpyxl).")
        return None
    path = os.path.join(tin_folder, STATUS_XLSX)

    # ---- rows already in the file from earlier runs
    old_heads, old_rows = [], []
    if os.path.exists(path):
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            if "Returns" in wb.sheetnames:
                rows = wb["Returns"].iter_rows(values_only=True)
                old_heads = [_clean(str(h)) if h is not None else ""
                             for h in next(rows, ())]
                for vals in rows:
                    d = {h: ("" if v is None else str(v))
                         for h, v in zip(old_heads, vals) if h}
                    if any(d.values()):
                        old_rows.append(d)
            wb.close()
        except Exception as err:
            log(f"   Could not read the old {STATUS_XLSX} "
                f"({type(err).__name__}) - starting it afresh.")
            old_heads, old_rows = [], []
    portal_heads = [h for h in old_heads
                    if h and h not in (BOT_TYPE, BOT_DL, BOT_CHK)]

    # ---- this run's rows
    new_rows = []
    for e in entries:
        d = {BOT_TYPE: e["type"]}
        for h, v in zip(e["heads"], e["cells"]):
            d[h] = v
            if h not in portal_heads:
                portal_heads.append(h)
        d[BOT_DL] = "Yes" if e["downloaded"] else "No"
        d[BOT_CHK] = checked_on
        new_rows.append(d)
    heads = [BOT_TYPE] + portal_heads + [BOT_DL, BOT_CHK]

    def find_head(*words):
        for h in portal_heads:
            hl = h.lower()
            if any(w in hl for w in words):
                return h
        return None
    key_head = find_head("return no", "return number", "ack", "reference")
    status_head = find_head("status")

    def key_of(d):
        if key_head and d.get(key_head, ""):
            return (d.get(BOT_TYPE, ""), d.get(key_head, ""))
        return (d.get(BOT_TYPE, ""),
                " | ".join(d.get(h, "") for h in portal_heads))

    merged = list(old_rows)
    where = {key_of(d): i for i, d in enumerate(merged)}
    for d in new_rows:
        k = key_of(d)
        if k in where:
            merged[where[k]] = d
        else:
            where[k] = len(merged)
            merged.append(d)

    # ---- Summary: return type x status
    def tally(rows):
        statuses, per_type = [], {}
        for d in rows:
            t = d.get(BOT_TYPE, "")
            st = (d.get(status_head, "") if status_head else "") or "(blank)"
            if status_head and st not in statuses:
                statuses.append(st)
            per_type.setdefault(t, {})
            per_type[t][st] = per_type[t].get(st, 0) + 1
        return statuses, per_type
    statuses, per_type = tally(merged)
    _, run_breakdown = tally(new_rows)

    # ---- build the workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append([BOT_TYPE] + statuses + ["Total"])
    for t, counts in per_type.items():
        ws.append([t] + [counts.get(st, 0) for st in statuses]
                  + [sum(counts.values())])
    ws2 = wb.create_sheet("Returns")
    ws2.append(heads)
    for d in merged:
        ws2.append([d.get(h, "") for h in heads])
    for sheet, freeze in ((ws, "A2"), (ws2, "B2")):
        for c in sheet[1]:
            c.font = Font(bold=True)
        sheet.freeze_panes = freeze
        sheet.auto_filter.ref = sheet.dimensions
        widths = {}
        for row in sheet.iter_rows():
            for c in row:
                if c.value is not None:
                    widths[c.column_letter] = max(
                        widths.get(c.column_letter, 0), len(str(c.value)))
        for col, w in widths.items():
            sheet.column_dimensions[col].width = min(max(w + 2, 8), 60)

    # ---- save via a temp name; if Excel has the file open, save beside it
    tmp = os.path.join(tin_folder, "~" + STATUS_XLSX)
    wb.save(tmp)
    try:
        os.replace(tmp, path)
    except OSError:
        alt = os.path.join(tin_folder, STATUS_XLSX.replace(
            ".xlsx", time.strftime("_%Y%m%d_%H%M.xlsx")))
        os.replace(tmp, alt)
        log(f"   {STATUS_XLSX} is open in another program - saved as "
            f"{os.path.basename(alt)} instead.")
        path = alt
    return {"path": path, "status_head": status_head,
            "breakdown": run_breakdown, "total": len(merged)}

# ======================================================================
# THE MAIN BOT
# ======================================================================
def run_bot(login_id, password, date_from, date_to, wanted_categories, log,
            patient=False, background=False, result=None):
    summary = []
    finished_ok = False
    if result is None:
        result = {}
    result["summary"] = summary          # filled in as categories finish
    result["ok"] = False
    result["folder"] = os.path.join(BASE_FOLDER, safe_name(login_id))
    _frame_memory.clear()     # fresh Chrome = fresh frame layout
    # rows for Returns_Status.xlsx (all return types, this client)
    status_rows = []
    status_written = False
    checked_on = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    # everything for this TIN goes into its own folder
    tin_folder = os.path.join(BASE_FOLDER, safe_name(login_id))

    # how long we are prepared to wait at each step of one return
    tab_limit = PATIENT_LIMIT_SECONDS if patient else TAB_WAIT_SECONDS
    page_limit = PATIENT_LIMIT_SECONDS if patient else PAGE_WAIT_SECONDS
    pdf_limit = PATIENT_LIMIT_SECONDS if patient else PDF_WAIT_SECONDS
    retries = 3 if patient else 1
    if patient:
        log("Mode: WAIT UNTIL IT DOWNLOADS - each return is given up to "
            f"{PATIENT_LIMIT_SECONDS // 60} minutes at every step, and "
            f"failed rows are retried {retries} times.")
    else:
        log("Mode: normal waits (tick 'Wait until it downloads' for slow "
            "returns).")

    log("Starting Chrome...")
    try:
        driver = make_driver()
    except Exception as err:
        log("ERROR: Could not start Chrome on this computer.")
        log(f"   Technical detail: {type(err).__name__}: {err}")
        log("   Checklist to fix this:")
        log("   1. Is GOOGLE CHROME installed on this computer? The bot")
        log("      needs real Chrome (not Edge or another browser).")
        log("   2. Is the INTERNET working right now? On its very first")
        log("      run on a new computer, the program downloads a small")
        log("      Chrome-driver in the background - that needs internet.")
        log("   3. Antivirus or company security may be blocking it -")
        log("      allow/whitelist this program and try again.")
        log("   4. Try right-click -> 'Run as administrator' once.")
        return
    try:
        # ---------------------------------------------------- 1. LOGIN
        log("Opening URA portal...")
        driver.get(LOGIN_URL)

        user_box = locate(driver, By.ID, "userName", 30)
        if user_box is None:
            log("ERROR: Login page did not load properly. Stopping.")
            return
        user_box.send_keys(login_id)
        pwd_box = driver.find_element(By.XPATH, "//input[@type='password']")
        pwd_box.send_keys(password)

        log(">>> NOW TYPE THE SECURITY CODE (CAPTCHA) IN CHROME AND CLICK 'Log In' <<<")

        # ------------------------------------------- 2. WAIT FOR LOGIN
        # The link is hunted by its TEXT through the page code, in every
        # window and every frame. Any 'Skip' button anywhere is clicked.
        history_link = None
        LINK_XPATH = ("//a[contains(translate(normalize-space(.),"
                      "'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                      "'abcdefghijklmnopqrstuvwxyz'),'return history')]")
        SKIP_XPATH = ("//button[normalize-space()='Skip']"
                      " | //input[@value='Skip']"
                      " | //*[@id='parishMsgA']//button"
                      " | //*[@id='NINBRNMsg']//button"
                      " | //*[@id='myModal']//button")
        for attempt in range(200):
            for handle in list(driver.window_handles):
                try:
                    driver.switch_to.window(handle)
                except Exception:
                    continue
                history_link = locate(driver, By.XPATH, LINK_XPATH, 1,
                                      require_visible=False)
                if history_link is not None:
                    break
                skip_btn = locate(driver, By.XPATH, SKIP_XPATH, 1)
                if skip_btn is not None:
                    try:
                        js_click(driver, skip_btn)
                        log("   (skipped a popup message)")
                    except Exception:
                        pass
            if history_link is not None:
                break
            if attempt % 10 == 0:
                # tell the user (and Claude!) what the bot can actually see
                marker = locate(driver, By.CSS_SELECTOR,
                                "#logoutFormISSP input", 1,
                                require_visible=False)
                if marker is not None:
                    log("   You appear LOGGED IN, but the 'return history' "
                        "menu is not on this screen - still searching...")
                else:
                    log("   ...waiting for you to log in...")

        if history_link is None:
            log("ERROR: Could not find 'View return history' anywhere.")
            log(f"   (Chrome had {len(driver.window_handles)} window(s) open; "
                f"last address: {driver.current_url})")
            log("Please copy this whole log to Claude.")
            return

        # ------------------------------------- 3. OPEN RETURN HISTORY
        log("Login detected! Opening 'View return history'...")
        if background:
            # Chrome had to be visible for the CAPTCHA; from here on it can
            # sit minimised in the taskbar while the returns are fetched.
            try:
                driver.minimize_window()
                log("Chrome minimised - the bot is working in the background. "
                    "Watch this window for progress.")
            except Exception:
                log("(Could not minimise Chrome - continuing with it visible.)")
        link_href = history_link.get_attribute("href")
        js_click(driver, history_link)

        search_btn = locate(driver, By.ID, "search", 25)
        if search_btn is None and link_href and "javascript" not in link_href.lower():
            # Plan B: go to the link's web address directly
            log("   Click did not open the page - trying its direct address...")
            driver.get(link_href)
            search_btn = locate(driver, By.ID, "search", 20)
        if search_btn is None:
            log("ERROR: The return-history search page did not load. Stopping.")
            return
        log("Return history page is open.")

        # ------------------------------------------ 4. READ CATEGORIES
        rtn_dropdown = locate(driver, By.ID, "rtnType", 15)
        categories = [(o.get_attribute("value"), o.text.strip())
                      for o in Select(rtn_dropdown).options
                      if o.get_attribute("value") and "select" not in o.text.lower()]
        log(f"Portal offers {len(categories)} return types: " +
            ", ".join(name for _, name in categories))

        # Keep only the categories the user TICKED (exact name match,
        # so 'Gaming Tax Tax Return' never accidentally also selects
        # 'Gaming Tax Tax Return - Weekly')
        if wanted_categories is not None:
            wanted_set = {w.lower() for w in wanted_categories}
            new_types = [name for _, name in categories
                         if name.lower() not in {c.lower() for c in ALL_CATEGORIES}]
            if new_types:
                log("NOTE: The portal now has new return type(s) not in the "
                    "tick list: " + ", ".join(new_types))
            categories = [(code, name) for code, name in categories
                          if name.lower() in wanted_set]
            if not categories:
                log("ERROR: None of your ticked categories matched the portal's "
                    "list above. Please tell Claude.")
                return
            log("Downloading only: " + ", ".join(name for _, name in categories))

        # ---- helper: download ONE return, surviving vanishing tabs ----
        def wait_for_new_tab(handles_before, limit, tag):
            """Wait up to `limit` seconds for the return to open in a new
            tab, reassuring the user every 30 s that we are still waiting."""
            t0 = time.time()
            last_note = t0
            while time.time() - t0 < limit:
                try:
                    extra = [h for h in driver.window_handles
                             if h not in handles_before]
                except WebDriverException:
                    extra = []
                if extra:
                    return extra[0]
                if time.time() - last_note >= 30:
                    log(f"   {tag} still waiting for the return to open "
                        f"({int(time.time() - t0)} s)...")
                    last_note = time.time()
                time.sleep(0.3)
            return None

        _LOW = "translate(%s,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz')"
        # 'Print with Schedules': by its id, or - in case URA ever renames
        # the id - by the words written on the button
        SCHED_XPATH = ("//*[@id='printWthSchid']"
                       " | //input[contains(" + _LOW % "@value"
                       + ",'print with sched')]"
                       " | //button[contains(" + _LOW % "normalize-space(.)"
                       + ",'print with sched')]"
                       " | //a[contains(" + _LOW % "normalize-space(.)"
                       + ",'print with sched')]")
        ANY_PRINT_XPATH = "//*[@id='printid'] | " + SCHED_XPATH

        def schedules_button_here():
            """'Print with Schedules' in the frame the browser is focused
            on right now (where a print button was just found)."""
            try:
                for el in driver.find_elements(By.XPATH, SCHED_XPATH):
                    try:
                        if el.is_displayed():
                            return el
                    except Exception:
                        continue
            except Exception:
                pass
            return None

        def wait_for_print_button(limit, tag):
            """The Print button only exists once the return has fully
            loaded, so waiting for it = waiting for the page.

            RULE: if the page has 'Print with Schedules' (VAT, PAYE,
            Withholding Tax ...) ALWAYS use it, so the schedules are in
            the PDF. Only when the page has no such button (Provisional,
            Non-Individual ...) use the plain 'Print'.

            The old code asked for "either button" in one search, and
            the browser hands back whichever comes FIRST on the page. On
            pages that carry both, plain 'Print' comes first, so the
            schedules were being left out."""
            t0 = time.time()
            last_note = t0
            while True:
                # one search for EITHER button = "has the page loaded?"
                btn = locate(driver, By.XPATH, ANY_PRINT_XPATH, 3)
                if btn is not None:
                    # locate() leaves the browser focused on the button's
                    # frame, so this look is instant. Allow half a second
                    # for the schedules button to be drawn next to 'Print',
                    # and up to 2 s if the page is still loading.
                    found_at = time.time()
                    grace_end = found_at + 2
                    while True:
                        sched = schedules_button_here()
                        if sched is not None:
                            log(f"   {tag} using 'Print with Schedules'")
                            return sched
                        try:
                            done = driver.execute_script(
                                "return document.readyState") == "complete"
                        except Exception:
                            done = True
                        if ((done and time.time() - found_at >= 0.5)
                                or time.time() >= grace_end):
                            break
                        time.sleep(0.2)
                    log(f"   {tag} no 'Print with Schedules' on this return "
                        "- using 'Print'")
                    return btn
                if time.time() - t0 >= limit:
                    return None
                if time.time() - last_note >= 30:
                    log(f"   {tag} return page still loading "
                        f"({int(time.time() - t0)} s)...")
                    last_note = time.time()

        UNBOX_JS = """
            var css = '.div1, [style*="overflow-X: scroll"], [style*="overflow-x: scroll"],'
                    + '[style*="overflow-x:scroll"], [style*="overflow-X:scroll"]'
                    + '{overflow:visible !important; width:auto !important;'
                    + ' margin-left:0 !important; margin-right:0 !important;}'
                    + '.div1 table, [style*="overflow-x"] table, [style*="overflow-X"] table'
                    + '{width:auto !important; min-width:100% !important;}'
                    + 'table.tab3 th, table.tab3 td {white-space:nowrap;}';
            var s = document.createElement('style');
            s.setAttribute('data-ura-unbox', '1');
            s.appendChild(document.createTextNode(css));
            (document.head || document.documentElement).appendChild(s);
            document.querySelectorAll('*').forEach(function (el) {
                var st = (el.getAttribute('style') || '').toLowerCase();
                if (st.indexOf('overflow') >= 0 && st.indexOf('scroll') >= 0) {
                    el.style.setProperty('overflow', 'visible', 'important');
                    el.style.setProperty('width', 'auto', 'important');
                }
            });
            return document.querySelectorAll('[data-ura-unbox]').length;
        """

        def unbox_scroll_boxes(depth=0):
            """Apply UNBOX_JS in the current document and in every frame
            beneath it (the portal nests frames up to 3 deep)."""
            try:
                driver.execute_script(UNBOX_JS)
            except Exception:
                pass
            if depth >= 3:
                return
            try:
                n = len(driver.find_elements(By.XPATH, "//iframe | //frame"))
            except Exception:
                return
            for i in range(n):
                try:
                    frames = driver.find_elements(By.XPATH, "//iframe | //frame")
                    if i >= len(frames):
                        break
                    driver.switch_to.frame(frames[i])
                except Exception:
                    continue
                unbox_scroll_boxes(depth + 1)
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    driver.switch_to.default_content()

        def save_web_copy(folder, row_label, tag):
            """Save the whole return page (every frame, styling, images)
            as ONE .mhtml file - exactly what is on screen, not a print
            layout. Uses Chrome's own 'capture snapshot' command."""
            try:
                driver.switch_to.default_content()
                # The portal hides wide schedules inside horizontally
                # scrolling boxes (overflow-X: scroll; width: 90%), so the
                # right-hand columns are clipped. Turn every such box in
                # every frame into a plain, full-width block before the
                # snapshot, so the saved page shows every column flat.
                unbox_scroll_boxes()
                time.sleep(0.5)          # let styles settle
                res = driver.execute_cdp_cmd("Page.captureSnapshot",
                                             {"format": "mhtml"})
                data = res.get("data", "") if isinstance(res, dict) else ""
                if not data:
                    log(f"   {tag} web-page copy came back empty - not saved")
                    return False
                target = unique_path(folder, row_label, ".mhtml")
                with open(target, "w", encoding="utf-8", newline="") as fh:
                    fh.write(data)
                log(f"   {tag} web page saved -> {os.path.basename(target)}")
                return True
            except Exception as err:
                log(f"   {tag} could not save web-page copy "
                    f"({type(err).__name__}: {err})")
                return False

        def download_one(view_link, folder, row_label, tag, save_html=False):
            """Open a return, print it, file the PDF. Excise returns
            (save_html=True) are saved as a .mhtml web copy INSTEAD of a
            PDF. Even if the tab closes itself
            mid-way, we still watch for the PDF, because the printing
            often completes anyway. Returns True on success."""
            handles_before = driver.window_handles
            started = time.time()
            js_click(driver, view_link)

            new_tab = wait_for_new_tab(handles_before, tab_limit, tag)
            if new_tab is None:
                log(f"   {tag} the return did not open in a new tab "
                    f"within {int(time.time() - started)} s")

            print_clicked_at = None
            html_saved = False
            try:
                if new_tab is not None:
                    driver.switch_to.window(new_tab)
                    print_btn = wait_for_print_button(page_limit, tag)
                    if print_btn is None:
                        log(f"   {tag} page did not finish loading "
                            f"({int(time.time() - started)} s)")
                    if save_html:
                        # Excise: web-page copy ONLY, no PDF at all
                        html_saved = save_web_copy(folder, row_label, tag)
                    elif print_btn is not None:
                        js_click(driver, print_btn)
                        print_clicked_at = time.time()
            except (NoSuchWindowException, WebDriverException):
                pass          # tab vanished - the PDF may still arrive

            if save_html:
                # nothing to print - tidy up and report
                try:
                    for h in list(driver.window_handles):
                        if h not in handles_before:
                            driver.switch_to.window(h)
                            driver.close()
                except Exception:
                    pass
                try:
                    driver.switch_to.window(handles_before[0])
                except Exception:
                    pass
                return html_saved

            # the PDF clock starts when Print was pressed, not at the
            # first click (that was the old bug that starved slow returns)
            clock = print_clicked_at or time.time()
            pdf = None
            last_note = clock
            while time.time() - clock < pdf_limit:
                pdf = newest_pdf(TEMP_PRINT_FOLDER, started - 2)
                if pdf:
                    break
                if time.time() - last_note >= 30:
                    log(f"   {tag} waiting for the PDF to finish "
                        f"({int(time.time() - clock)} s)...")
                    last_note = time.time()
                time.sleep(0.3)

            # tidy up: close any extra windows, go back to the main one
            try:
                for h in list(driver.window_handles):
                    if h not in handles_before:
                        driver.switch_to.window(h)
                        driver.close()
            except Exception:
                pass
            try:
                driver.switch_to.window(handles_before[0])
            except Exception:
                pass

            if pdf:
                target = unique_path(folder, row_label, ".pdf")
                shutil.move(pdf, target)
                log(f"   {tag} saved -> {os.path.basename(target)}")
                return True
            return False
        # ----------------------------------------------------------------

        # -------------------------------------------- 5. LOOP CATEGORIES
        for code, name in categories:
            folder = os.path.join(tin_folder, safe_name(name))
            os.makedirs(folder, exist_ok=True)
            save_html = name.strip().lower() in {c.lower() for c in HTML_CATEGORIES}
            log(f"\n=== {name} ===" +
                ("   (saved as web page .mhtml, no PDF)" if save_html else ""))

            if not fill_dates(driver, date_from, date_to, log):
                log("   Skipping this category (no date boxes).")
                continue
            dd = locate(driver, By.ID, "rtnType", 10)
            Select(dd).select_by_value(code)
            # Excise / VAT / PAYE / WHT show a Headquarter/Both box (preset
            # to Headquarter) - switch it to Both so every branch is listed
            if code in BRANCH_CODES:
                select_both_branches(driver, log)
            btn = locate(driver, By.ID, "search", 10)
            js_click(driver, btn)
            # continue the moment the results area is there (no fixed pause)

            # ---------------------------------------- 6. LOOP RESULT ROWS
            # The results table shows only ONE PAGE at a time. We read
            # 'Total Records : N' as our checklist, download every row on
            # the page, then click 'Next' until all pages are done.
            table_area = locate(driver, By.ID, "returnHistory", 15)
            if table_area is None:
                log("   No results area appeared - skipping.")
                summary.append((name, 0))
                continue

            total_records = None
            try:
                frame_text = driver.find_element(By.TAG_NAME, "body").text
            except Exception:
                frame_text = table_area.text
            m = re.search(r"Total\s*Records?\s*:\s*(\d+)", frame_text)
            if m:
                total_records = int(m.group(1))
                log(f"   Portal reports {total_records} record(s) in total.")
            try:
                headings = read_history_headings(driver)
            except Exception:
                headings = [f"Column {i + 1}" for i in range(10)]

            downloaded = 0
            indexed = 0
            page = 1
            while True:
                locate(driver, By.ID, "returnHistory", 10)
                rows_now = driver.find_elements(
                    By.XPATH, "//*[@id='returnHistory']//tbody/tr[.//td[11]/a]")
                if not rows_now:
                    if page == 1:
                        log("   No returns found for this period.")
                    break
                log(f"   Page {page}: {len(rows_now)} return(s) on this page.")

                # copy the rows for the status sheet BEFORE anything is
                # clicked (the table redraws after every return)
                page_entries = {}
                for rn, r in enumerate(rows_now):
                    try:
                        cells = read_row_cells(r)
                    except Exception:
                        cells = [""] * 10
                    entry = {"type": name, "heads": headings,
                             "cells": cells, "downloaded": False}
                    page_entries[rn] = entry
                    status_rows.append(entry)
                    indexed += 1

                failed_rows = []
                for row_number in range(len(rows_now)):
                    try:
                        locate(driver, By.ID, "returnHistory", 10)
                        rows_now = driver.find_elements(
                            By.XPATH,
                            "//*[@id='returnHistory']//tbody/tr[.//td[11]/a]")
                        if row_number >= len(rows_now):
                            break
                        row = rows_now[row_number]
                        row_label = safe_name(" ".join(row.text.split())[:60]) \
                            or f"p{page}_return_{row_number + 1}"
                        view_link = row.find_element(By.XPATH, ".//td[11]/a")
                        if download_one(view_link, folder, row_label,
                                        f"[page {page}, row {row_number + 1}]",
                                        save_html):
                            downloaded += 1
                            if row_number in page_entries:
                                page_entries[row_number]["downloaded"] = True
                        else:
                            log(f"   [page {page}, row {row_number + 1}] "
                                "not saved - will retry at end of page")
                            failed_rows.append(row_number)
                    except Exception as err:
                        log(f"   [page {page}, row {row_number + 1}] problem "
                            f"({type(err).__name__}) - will retry at end of page")
                        failed_rows.append(row_number)
                        try:
                            while len(driver.window_handles) > 1:
                                driver.switch_to.window(driver.window_handles[-1])
                                driver.close()
                            driver.switch_to.window(driver.window_handles[0])
                        except Exception:
                            pass

                # RETRY every row that failed on this page (1 round in
                # normal mode, 3 rounds in "wait until it downloads" mode)
                for retry_round in range(1, retries + 1):
                    if not failed_rows:
                        break
                    still_failed = []
                    log(f"   Retry round {retry_round} of {retries} for "
                        f"{len(failed_rows)} row(s) on page {page}...")
                    for row_number in failed_rows:
                        label = (f"[RETRY {retry_round} page {page}, "
                                 f"row {row_number + 1}]")
                        try:
                            locate(driver, By.ID, "returnHistory", 10)
                            rows_retry = driver.find_elements(
                                By.XPATH,
                                "//*[@id='returnHistory']//tbody/tr[.//td[11]/a]")
                            if row_number >= len(rows_retry):
                                continue
                            row = rows_retry[row_number]
                            row_label = safe_name(" ".join(row.text.split())[:60]) \
                                or f"p{page}_return_{row_number + 1}"
                            view_link = row.find_element(By.XPATH, ".//td[11]/a")
                            if download_one(view_link, folder, row_label,
                                            label, save_html):
                                downloaded += 1
                                if row_number in page_entries:
                                    page_entries[row_number]["downloaded"] = True
                            else:
                                log(f"   {label} still no PDF")
                                still_failed.append(row_number)
                        except Exception as err:
                            log(f"   {label} problem again "
                                f"({type(err).__name__})")
                            still_failed.append(row_number)
                            try:
                                while len(driver.window_handles) > 1:
                                    driver.switch_to.window(driver.window_handles[-1])
                                    driver.close()
                                driver.switch_to.window(driver.window_handles[0])
                            except Exception:
                                pass
                    failed_rows = still_failed
                    if failed_rows and retry_round < retries:
                        time.sleep(5)           # give the portal a breather
                for row_number in failed_rows:
                    log(f"   *** [page {page}, row {row_number + 1}] could not "
                        "be downloaded after all retries - skipped ***")

                # Are we on the last page? Read URA's own page counter
                # ('Current Page : 2 / 2') and stop without a futile click.
                locate(driver, By.ID, "returnHistory", 10)
                try:
                    page_text = driver.find_element(By.TAG_NAME, "body").text
                except Exception:
                    page_text = ""
                pinfo = re.search(r"Current\s*Page\s*:\s*(\d+)\s*/\s*(\d+)",
                                  page_text)
                if pinfo and int(pinfo.group(1)) >= int(pinfo.group(2)):
                    log(f"   Last page reached "
                        f"({pinfo.group(1)}/{pinfo.group(2)}).")
                    break

                before_rows = driver.find_elements(
                    By.XPATH, "//*[@id='returnHistory']//tbody/tr[.//td[11]/a]")
                before_sig = before_rows[0].text if before_rows else ""

                # Prefer a real link/button named 'Next'. A plain table cell
                # that merely CONTAINS the link gives a "dead click", so if
                # only a wrapper matches, we click the link INSIDE it.
                next_btn = locate(
                    driver, By.XPATH,
                    "//a[contains(normalize-space(.),'Next')"
                    " and string-length(normalize-space(.)) <= 10]"
                    " | //input[@value='Next']"
                    " | //button[normalize-space(.)='Next']",
                    5, require_visible=False)
                if next_btn is None:
                    wrapper = locate(
                        driver, By.XPATH,
                        "//*[self::td or self::span or self::div]"
                        "[contains(normalize-space(.),'Next')"
                        " and string-length(normalize-space(.)) <= 10]",
                        3, require_visible=False)
                    if wrapper is not None:
                        inner = wrapper.find_elements(
                            By.XPATH, ".//a | .//input | .//button")
                        next_btn = inner[0] if inner else wrapper
                if next_btn is None:
                    log("   No 'Next' button found - treating this as the last page.")
                    break
                log(f"   Turning to page {page + 1}...")
                js_click(driver, next_btn)

                changed = False
                for _ in range(40):             # up to ~12s for page to load
                    time.sleep(0.3)
                    locate(driver, By.ID, "returnHistory", 5)
                    rows_check = driver.find_elements(
                        By.XPATH,
                        "//*[@id='returnHistory']//tbody/tr[.//td[11]/a]")
                    sig = rows_check[0].text if rows_check else ""
                    if sig and sig != before_sig:
                        changed = True
                        break
                if not changed:
                    log("   'Next' was clicked but the table did not change - "
                        "treating this as the last page.")
                    break
                page += 1

            # The safeguard: did we get everything the portal promised?
            if total_records is not None:
                if downloaded < total_records:
                    log(f"   *** WARNING: portal reported {total_records} "
                        f"record(s) but only {downloaded} file(s) were saved. "
                        "Some returns may be missing - please re-check this "
                        "category. ***")
                else:
                    log(f"   All {total_records} record(s) downloaded - complete.")

            summary.append((name, downloaded))
            if indexed:
                log(f"   Indexed {indexed} row(s) for the status sheet.")

        # ------------------------------------------------- 7. SUMMARY
        log("\n==================== FINISHED ====================")
        for name, count in summary:
            log(f"   {name:35s} {count} file(s)")
        log(f"\nAll files are in: {tin_folder}")

        # the Excel status sheet (all return types of this client)
        try:
            info = write_status_sheet(tin_folder, status_rows, checked_on, log)
            status_written = True
            if info:
                log(f"Status sheet: {os.path.basename(info['path'])}  "
                    f"({len(status_rows)} return(s) this run, "
                    f"{info['total']} listed in all - see the Summary tab)")
                if info["status_head"]:
                    for t, counts in info["breakdown"].items():
                        parts = ", ".join(f"{st} {n}" for st, n in counts.items())
                        log(f"   {t:35s} {info['status_head']}: {parts}")
                else:
                    log("   No status column found in the Return History "
                        "table - the Summary tab shows counts only.")
        except Exception as err:
            log(f"ERROR: could not write {STATUS_XLSX} "
                f"({type(err).__name__}: {err}) - the downloads are unaffected.")

        # Log out - the visible Logout button inside form 'logoutFormISSP'
        # (NOT the form's hidden bookkeeping inputs, which caused dead clicks)
        log("Logging out...")
        logout_btn = locate(
            driver, By.XPATH,
            "//form[@id='logoutFormISSP']//input[@value='Logout' or "
            "@type='submit' or @type='button']"
            " | //*[@id='logoutFormISSP']/table/tbody/tr[1]/td/input",
            6)
        if logout_btn is None:   # backup: text-based hunt anywhere on the page
            logout_btn = locate(
                driver, By.XPATH,
                "//*[self::a or self::button]"
                "[normalize-space(text())='Logout' or normalize-space(text())='Log Out']"
                " | //input[@value='Logout' or @value='Log Out']",
                4, require_visible=False)

        clicked = False
        if logout_btn is not None:
            try:
                js_click(driver, logout_btn)
                clicked = True
            except Exception:
                pass

        def looks_logged_out():
            """Logged out if the login box is back OR the logout form is gone."""
            if locate(driver, By.ID, "userName", 2) is not None:
                return True
            return locate(driver, By.CSS_SELECTOR, "#logoutFormISSP input", 2,
                          require_visible=False) is None

        time.sleep(2)
        confirmed = clicked and looks_logged_out()
        if not confirmed:
            form = locate(driver, By.ID, "logoutFormISSP", 3, require_visible=False)
            if form is not None:
                try:
                    driver.execute_script("arguments[0].submit();", form)
                    clicked = True
                except Exception:
                    pass
                time.sleep(2)
                confirmed = looks_logged_out()

        if confirmed:
            log("Logged out of the URA portal - confirmed.")
        elif clicked:
            log("Logout was clicked - please glance at Chrome to confirm.")
        else:
            log("Could not find the Logout button - please log out manually.")

        try:
            os.rmdir(TEMP_PRINT_FOLDER)
        except OSError:
            pass
        finished_ok = True
        result["ok"] = True

    except Exception as fatal:
        log(f"\nFATAL ERROR: {type(fatal).__name__}: {fatal}")
    finally:
        if status_rows and not status_written:
            # run stopped early - still keep what was collected
            try:
                info = write_status_sheet(tin_folder, status_rows, checked_on, log)
                if info:
                    log(f"Status sheet saved with what was collected so far: "
                        f"{os.path.basename(info['path'])}")
            except Exception as err:
                log(f"Could not write {STATUS_XLSX} ({type(err).__name__}).")
        if finished_ok:
            try:
                driver.quit()
                log("Chrome closed.")
            except Exception:
                pass
        else:
            if background:
                try:
                    driver.switch_to.window(driver.window_handles[0])
                    driver.maximize_window()
                except Exception:
                    pass
            log("Chrome was left open so you can see what happened.")


# ======================================================================
# LOOK AND FEEL
# ======================================================================
# Ink navy, paper grey and a stamp teal: the colours of a filed return.
INK        = "#12283D"
INK_DEEP   = "#0B1B2B"
PAPER      = "#EEF1F5"
CARD       = "#FFFFFF"
LINE       = "#D5DCE5"
TEXT       = "#16232F"
MUTED      = "#6B7A8C"
TEAL       = "#0E6E70"
TEAL_DARK  = "#095052"
TEAL_LIT   = "#12888A"
GREY_BTN   = "#E3E8EF"
GREY_DARK  = "#C2CBD6"
STAMP_RED  = "#B5342B"
LEDGER_GRN = "#2E7D52"
AMBER_BG   = "#FDF3E0"
AMBER_LINE = "#E8C88A"
AMBER_TEXT = "#7A5314"

FONT       = ("Segoe UI", 10)
FONT_SMALL = ("Segoe UI", 9)
FONT_MED   = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 19, "bold")
FONT_STAT  = ("Segoe UI", 20, "bold")
FONT_LOG   = ("Consolas", 9)


def resource_path(name):
    """Find a file that travels with the program. PyInstaller unpacks
    into a temporary folder at run time, so the icon does not simply
    sit next to the .exe."""
    base = getattr(sys, "_MEIPASS", None)
    if base is None:
        try:
            base = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            base = os.getcwd()
    return os.path.join(base, name)


def rounded(canvas, x1, y1, x2, y2, r, **kw):
    """Draw a rounded rectangle on a canvas."""
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
           x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
           x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(pts, smooth=True, **kw)


class Button3D(tk.Canvas):
    """A raised button that physically compresses when pressed: the face
    sits above a hard shadow and drops onto it on click."""

    DEPTH = 5

    def __init__(self, parent, text, command, width=190, height=44,
                 face=TEAL, face_lit=TEAL_LIT, shadow=TEAL_DARK,
                 fg="#FFFFFF", bg=CARD, font=FONT_MED, radius=9):
        super().__init__(parent, width=width, height=height + self.DEPTH,
                         bg=bg, highlightthickness=0, bd=0)
        self._bw, self._bh, self._r = width, height, radius
        self._face, self._lit, self._shadow = face, face_lit, shadow
        self._fg, self._font, self._text = fg, font, text
        self._command = command
        self._enabled = True
        self._down = False
        self._hover = False
        self._draw()
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)

    def _draw(self):
        self.delete("all")
        d = self.DEPTH
        offset = d - 1 if self._down else 0
        if self._enabled:
            face = self._lit if (self._hover and not self._down) else self._face
            shadow, fg = self._shadow, self._fg
        else:
            face, shadow, fg = GREY_BTN, GREY_DARK, MUTED
        rounded(self, 1, d, self._bw - 1, self._bh + d, self._r,
                fill=shadow, outline="")
        rounded(self, 1, offset, self._bw - 1, self._bh + offset, self._r,
                fill=face, outline="")
        self.create_text(self._bw / 2, self._bh / 2 + offset,
                         text=self._text, fill=fg, font=self._font)

    def _on_enter(self, _):
        self._hover = True
        self._draw()

    def _on_leave(self, _):
        self._hover = self._down = False
        self._draw()

    def _on_press(self, _):
        if not self._enabled:
            return
        self._down = True
        self._draw()

    def _on_release(self, _):
        if not self._enabled or not self._down:
            return
        self._down = False
        self._draw()
        if self._command:
            self._command()

    def set_text(self, text):
        self._text = text
        self._draw()

    def configure(self, **kw):          # keeps .configure(state=...) working
        if "state" in kw:
            self._enabled = kw.pop("state") != "disabled"
            self._down = self._hover = False
            super().configure(cursor="hand2" if self._enabled else "")
            self._draw()
        if kw:
            super().configure(**kw)

    config = configure


class CheckBox(tk.Canvas):
    """A tick box drawn by hand, so it matches everything else."""

    def __init__(self, parent, text, variable, bg=CARD, width=None,
                 full_text=None):
        self.var = variable
        self._text = text
        self._full = full_text or text
        if width is None:
            width = 26 + 7 * len(text)
        super().__init__(parent, width=width, height=24, bg=bg,
                         highlightthickness=0, bd=0, cursor="hand2")
        self._draw()
        self.bind("<Button-1>", self._toggle)

    def _draw(self):
        self.delete("all")
        on = bool(self.var.get())
        rounded(self, 2, 4, 18, 20, 4,
                fill=TEAL if on else "#FFFFFF",
                outline=TEAL if on else LINE)
        if on:
            self.create_line(6, 12, 9, 15, 14, 8, fill="#FFFFFF", width=2,
                             capstyle="round", joinstyle="round")
        self.create_text(26, 12, text=self._text, anchor="w",
                         fill=TEXT if on else MUTED, font=FONT_SMALL)

    def _toggle(self, _=None):
        self.var.set(not self.var.get())
        self._draw()

    def refresh(self):
        self._draw()


class ProgressBar(tk.Canvas):
    """A slim bar. Because the portal never says how many returns exist
    before it finds them, this pulses rather than pretending to know a
    percentage."""

    def __init__(self, parent, width=880, height=10, bg=PAPER):
        super().__init__(parent, width=width, height=height, bg=bg,
                         highlightthickness=0, bd=0)
        self._bw, self._bh = width, height
        self._pulsing = False
        self._pos = 0.0
        self.bind("<Configure>", self._resized)
        self._draw()

    def _resized(self, event):
        self._bw = event.width
        self._draw()

    def _draw(self):
        self.delete("all")
        r = self._bh / 2
        rounded(self, 0, 0, self._bw, self._bh, r, fill="#E1E7EF", outline="")
        if self._pulsing:
            seg = 0.22
            left = self._pos * (1 + seg) - seg
            x1 = max(0.0, left) * self._bw
            x2 = min(1.0, left + seg) * self._bw
            if x2 - x1 > self._bh:
                rounded(self, x1, 0, x2, self._bh, r, fill=TEAL, outline="")

    def pulse_on(self):
        if self._pulsing:
            return
        self._pulsing = True
        self._step()

    def _step(self):
        if not self._pulsing:
            return
        self._pos = (self._pos + 0.015) % 1.0
        self._draw()
        self.after(35, self._step)

    def pulse_off(self):
        self._pulsing = False
        self._draw()


class Stat(tk.Frame):
    """One live counter: a number with a quiet label beneath it."""

    def __init__(self, parent, label, colour=TEXT, bg=PAPER):
        super().__init__(parent, bg=bg)
        self.value = tk.StringVar(value="0")
        tk.Label(self, textvariable=self.value, font=FONT_STAT,
                 fg=colour, bg=bg).pack(anchor="w")
        tk.Label(self, text=label, font=FONT_SMALL, fg=MUTED,
                 bg=bg).pack(anchor="w")


def field(parent, width=26, show=None):
    """A text box with a flat border and a focus ring."""
    holder = tk.Frame(parent, bg=LINE, padx=1, pady=1)
    entry = tk.Entry(holder, width=width, font=FONT, relief="flat",
                     bg="#FFFFFF", fg=TEXT, insertbackground=TEAL,
                     highlightthickness=0, bd=6, show=show)
    entry.pack()
    entry.bind("<FocusIn>", lambda e: holder.configure(bg=TEAL))
    entry.bind("<FocusOut>", lambda e: holder.configure(bg=LINE))
    return holder, entry


# ======================================================================
# THE WINDOW
# ======================================================================
def main():
    root = tk.Tk()
    root.title("URA Returns Downloader")
    root.geometry("1000x960")
    root.minsize(940, 860)
    root.configure(bg=PAPER)
    try:
        root.iconbitmap(resource_path("ura_icon.ico"))
    except Exception:
        pass          # a missing icon is not worth failing over

    # ---------------- header band ----------------
    header = tk.Frame(root, bg=INK, height=86)
    header.pack(fill="x")
    header.pack_propagate(False)
    head_in = tk.Frame(header, bg=INK)
    head_in.pack(fill="both", expand=True, padx=26, pady=14)
    tk.Label(head_in, text="URA returns downloader", font=FONT_TITLE,
             fg="#FFFFFF", bg=INK).pack(anchor="w")
    tk.Label(head_in, text="Signs in to ura.go.ug and files every return "
                           "you have submitted as a PDF",
             font=FONT_SMALL, fg="#93A7BC", bg=INK).pack(anchor="w")

    tear = tk.Canvas(root, height=6, bg=PAPER, highlightthickness=0, bd=0)
    tear.pack(fill="x")
    tear.bind("<Configure>", lambda e: (
        tear.delete("all"),
        tear.create_line(0, 3, e.width, 3, fill=LINE, dash=(3, 4))))

    body = tk.Frame(root, bg=PAPER)
    body.pack(fill="both", expand=True, padx=22, pady=(10, 16))

    # ---------------- clients + period card ----------------
    card = tk.Frame(body, bg=CARD, highlightbackground=LINE,
                    highlightthickness=1)
    card.pack(fill="x")
    pad = tk.Frame(card, bg=CARD)
    pad.pack(fill="x", padx=22, pady=16)
    pad.columnconfigure(0, weight=1)

    # ----- left: saved clients with tick boxes -----
    clients = load_clients()
    ticked_logins = set()

    left = tk.Frame(pad, bg=CARD)
    left.grid(row=0, column=0, sticky="nwe", padx=(0, 30))
    head = tk.Frame(left, bg=CARD)
    head.pack(fill="x")
    tk.Label(head, text="Clients", font=FONT_MED, fg=TEXT, bg=CARD).pack(side="left")
    client_count = tk.Label(head, text="", font=FONT_SMALL, fg=MUTED, bg=CARD)
    client_count.pack(side="left", padx=(10, 0))
    tk.Label(head, text="tick the ones to download - you type the security "
                        "code once per client",
             font=FONT_SMALL, fg=MUTED, bg=CARD).pack(side="right")

    style = ttk.Style()
    style.configure("Clients.Treeview", font=FONT, rowheight=24,
                    background="#FFFFFF", fieldbackground="#FFFFFF")
    style.configure("Clients.Treeview.Heading", font=FONT_SMALL)
    tree_wrap = tk.Frame(left, bg=LINE, padx=1, pady=1)
    tree_wrap.pack(fill="x", pady=(6, 0))
    tree = ttk.Treeview(tree_wrap, columns=("tick", "name", "login"),
                        show="headings", height=4, style="Clients.Treeview",
                        selectmode="browse")
    tree.heading("tick", text="")
    tree.heading("name", text="Client name")
    tree.heading("login", text="Login ID (TIN)")
    tree.column("tick", width=34, anchor="center", stretch=False)
    tree.column("name", width=330, anchor="w")
    tree.column("login", width=160, anchor="w")
    tree.pack(side="left", fill="x", expand=True)
    tree_scroll = ttk.Scrollbar(tree_wrap, command=tree.yview)
    tree.configure(yscrollcommand=tree_scroll.set)
    tree_scroll.pack(side="right", fill="y")

    def refresh_clients():
        tree.delete(*tree.get_children())
        for i, c in enumerate(clients):
            mark = "\u2611" if c["login"] in ticked_logins else "\u2610"
            tree.insert("", "end", iid=str(i), values=(mark, c["name"], c["login"]))
        n = len(ticked_logins & {c["login"] for c in clients})
        client_count.configure(text=f"{n} of {len(clients)} ticked")

    def on_tree_click(event):
        iid = tree.identify_row(event.y)
        if not iid:
            return
        login = clients[int(iid)]["login"]
        if login in ticked_logins:
            ticked_logins.discard(login)
        else:
            ticked_logins.add(login)
        refresh_clients()
    tree.bind("<Button-1>", on_tree_click)

    def client_dialog(existing=None):
        """Small window to add or edit one client. Returns dict or None."""
        dlg = tk.Toplevel(root)
        dlg.title("Edit client" if existing else "Add client")
        dlg.configure(bg=CARD)
        dlg.transient(root)
        dlg.grab_set()
        dlg.resizable(False, False)
        box_ = tk.Frame(dlg, bg=CARD)
        box_.pack(padx=22, pady=18)
        entries = {}
        for r, (key, caption, show) in enumerate([
                ("name", "Client name", None),
                ("login", "Login ID (TIN)", None),
                ("password", "Password", "*")]):
            tk.Label(box_, text=caption, font=FONT_MED, fg=TEXT,
                     bg=CARD).grid(row=r, column=0, sticky="w", pady=(0, 8),
                                   padx=(0, 14))
            holder, ent = field(box_, width=32, show=show)
            holder.grid(row=r, column=1, sticky="w", pady=(0, 8))
            if existing:
                ent.insert(0, existing.get(key, ""))
            entries[key] = ent
        out = {}

        def ok():
            vals = {k: e.get().strip() if k != "password" else e.get()
                    for k, e in entries.items()}
            if not (vals["name"] and vals["login"] and vals["password"]):
                messagebox.showwarning("Missing info",
                                       "Please fill in all three boxes.",
                                       parent=dlg)
                return
            out.update(vals)
            dlg.destroy()
        row_ = tk.Frame(box_, bg=CARD)
        row_.grid(row=3, column=0, columnspan=2, sticky="e", pady=(6, 0))
        Button3D(row_, "Cancel", dlg.destroy, width=90, height=32,
                 face=GREY_BTN, face_lit="#EDF1F6", shadow=GREY_DARK,
                 fg=TEXT, font=FONT_SMALL).pack(side="right", padx=(8, 0))
        Button3D(row_, "Save", ok, width=90, height=32,
                 font=FONT_SMALL).pack(side="right")
        entries["name"].focus_set()
        dlg.bind("<Return>", lambda e: ok())
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        root.wait_window(dlg)
        return out or None

    def add_client():
        c = client_dialog()
        if c:
            clients.append(c)
            ticked_logins.add(c["login"])
            save_clients(clients)
            refresh_clients()

    def selected_index():
        sel = tree.selection()
        return int(sel[0]) if sel else None

    def edit_client():
        i = selected_index()
        if i is None:
            messagebox.showinfo("Edit client", "Click a client in the list first.")
            return
        c = client_dialog(clients[i])
        if c:
            clients[i] = c
            save_clients(clients)
            refresh_clients()

    def remove_client():
        i = selected_index()
        if i is None:
            messagebox.showinfo("Remove client", "Click a client in the list first.")
            return
        if messagebox.askyesno("Remove client",
                               f"Remove {clients[i]['name']} from the list?"):
            ticked_logins.discard(clients[i]["login"])
            del clients[i]
            save_clients(clients)
            refresh_clients()

    def tick_all_clients():
        ticked_logins.update(c["login"] for c in clients)
        refresh_clients()

    def untick_all_clients():
        ticked_logins.clear()
        refresh_clients()

    cbtns = tk.Frame(left, bg=CARD)
    cbtns.pack(anchor="w", pady=(10, 0))
    for label_, fn in (("Add", add_client), ("Edit", edit_client),
                       ("Remove", remove_client), ("Tick all", tick_all_clients),
                       ("Untick all", untick_all_clients)):
        Button3D(cbtns, label_, fn, width=88, height=30, face=GREY_BTN,
                 face_lit="#EDF1F6", shadow=GREY_DARK, fg=TEXT,
                 font=FONT_SMALL).pack(side="left", padx=(0, 8))
    refresh_clients()

    # ----- right: the period -----
    right = tk.Frame(pad, bg=CARD)
    right.grid(row=0, column=1, sticky="ne")

    def labelled(parent, caption, hint=None):
        wrap = tk.Frame(parent, bg=CARD)
        wrap.pack(anchor="w", pady=(0, 10))
        line = tk.Frame(wrap, bg=CARD)
        line.pack(anchor="w")
        tk.Label(line, text=caption, font=FONT_MED, fg=TEXT,
                 bg=CARD).pack(side="left")
        if hint:
            tk.Label(line, text="  " + hint, font=FONT_SMALL, fg=MUTED,
                     bg=CARD).pack(side="left")
        holder, entry = field(wrap)
        holder.pack(anchor="w", pady=(4, 0))
        return entry

    e_from = labelled(right, "From date", "dd/mm/yyyy")
    e_to = labelled(right, "To date", "dd/mm/yyyy")

    # the one thing a person must do by hand, said plainly
    note = tk.Frame(pad, bg=AMBER_BG, highlightbackground=AMBER_LINE,
                    highlightthickness=1)
    note.grid(row=2, column=0, columnspan=2, sticky="we", pady=(4, 0))
    tk.Label(note, text="Chrome opens for each ticked client in turn. Type the "
                        "security code there and click Log In — the bot takes "
                        "over from that point and moves to the next client "
                        "when done.",
             font=FONT_SMALL, fg=AMBER_TEXT, bg=AMBER_BG, wraplength=880,
             justify="left", anchor="w").pack(fill="x", padx=12, pady=8)

    # slow-portal switch: keep waiting for each return instead of skipping
    patient_var = tk.BooleanVar(value=True)
    patient_row = tk.Frame(pad, bg=CARD)
    patient_row.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))
    CheckBox(patient_row,
             "Wait until each return downloads (slower, never skips a slow return)",
             patient_var, width=520).pack(side="left")
    tk.Label(patient_row,
             text="Excise returns are saved as a web page (.mhtml), not PDF.",
             font=FONT_SMALL, fg=MUTED, bg=CARD).pack(side="left", padx=(6, 0))

    background_var = tk.BooleanVar(value=True)
    bg_row = tk.Frame(pad, bg=CARD)
    bg_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))
    CheckBox(bg_row,
             "Run in background - minimise Chrome after you log in",
             background_var, width=520).pack(side="left")
    tk.Label(bg_row,
             text="Chrome stays visible until you have typed the security code.",
             font=FONT_SMALL, fg=MUTED, bg=CARD).pack(side="left", padx=(6, 0))

    # ---------------- return types card ----------------
    cats = tk.Frame(body, bg=CARD, highlightbackground=LINE,
                    highlightthickness=1)
    cats.pack(fill="x", pady=(12, 0))
    cpad = tk.Frame(cats, bg=CARD)
    cpad.pack(fill="x", padx=22, pady=16)

    top_line = tk.Frame(cpad, bg=CARD)
    top_line.pack(fill="x")
    tk.Label(top_line, text="Return types to download", font=FONT_MED,
             fg=TEXT, bg=CARD).pack(side="left")
    count_lbl = tk.Label(top_line, text="", font=FONT_SMALL, fg=MUTED,
                         bg=CARD)
    count_lbl.pack(side="left", padx=(10, 0))

    grid = tk.Frame(cpad, bg=CARD)
    grid.pack(fill="x", pady=(10, 0))

    cat_vars, cat_boxes = {}, []
    per_col = 6
    for i, cat in enumerate(ALL_CATEGORIES):
        var = tk.BooleanVar(value=(cat in DEFAULT_TICKED))
        cat_vars[cat] = var
        shown = cat if len(cat) <= 42 else cat[:39] + "..."
        cb = CheckBox(grid, shown, var, width=300, full_text=cat)
        cb.grid(row=i % per_col, column=i // per_col, sticky="w",
                padx=(0, 10), pady=1)
        cat_boxes.append(cb)

    def refresh_count():
        n = sum(1 for v in cat_vars.values() if v.get())
        count_lbl.configure(text=f"{n} of {len(ALL_CATEGORIES)} ticked")

    for cb in cat_boxes:
        original = cb._toggle

        def wrapped(event, fn=original):
            fn(event)
            refresh_count()
        cb.bind("<Button-1>", wrapped)

    def select_all():
        for v in cat_vars.values():
            v.set(True)
        for cb in cat_boxes:
            cb.refresh()
        refresh_count()

    def clear_all():
        for v in cat_vars.values():
            v.set(False)
        for cb in cat_boxes:
            cb.refresh()
        refresh_count()

    small_btns = tk.Frame(cpad, bg=CARD)
    small_btns.pack(anchor="w", pady=(12, 0))
    Button3D(small_btns, "Select all", select_all, width=100, height=32,
             face=GREY_BTN, face_lit="#EDF1F6", shadow=GREY_DARK,
             fg=TEXT, font=FONT_SMALL).pack(side="left")
    Button3D(small_btns, "Clear all", clear_all, width=100, height=32,
             face=GREY_BTN, face_lit="#EDF1F6", shadow=GREY_DARK,
             fg=TEXT, font=FONT_SMALL).pack(side="left", padx=8)
    refresh_count()

    # ---------------- action row ----------------
    action = tk.Frame(body, bg=PAPER)
    action.pack(fill="x", pady=(14, 0))
    btn = Button3D(action, "Fetch returns", lambda: start(), width=190,
                   height=46, bg=PAPER)
    btn.pack(side="left")
    Button3D(action, "History", lambda: show_history(), width=110, height=46,
             face=GREY_BTN, face_lit="#EDF1F6", shadow=GREY_DARK, fg=TEXT,
             bg=PAPER).pack(side="left", padx=(10, 0))

    stats = tk.Frame(action, bg=PAPER)
    stats.pack(side="right")
    stat_saved = Stat(stats, "saved", LEDGER_GRN)
    stat_saved.pack(side="left", padx=(0, 30))
    stat_issue = Stat(stats, "problems", STAMP_RED)
    stat_issue.pack(side="left")

    # ---------------- progress ----------------
    prog_wrap = tk.Frame(body, bg=PAPER)
    prog_wrap.pack(fill="x", pady=(14, 0))
    status = tk.Label(prog_wrap, text="Fill in the boxes and press fetch.",
                      font=FONT_SMALL, fg=MUTED, bg=PAPER, anchor="w")
    status.pack(fill="x", pady=(0, 6))
    bar = ProgressBar(prog_wrap)
    bar.pack(fill="x")

    # ---------------- activity panel ----------------
    panel = tk.Frame(body, bg=INK_DEEP)
    panel.pack(fill="both", expand=True, pady=(14, 0))
    tk.Label(panel, text="Activity", font=FONT_SMALL, fg="#7B90A6",
             bg=INK_DEEP, anchor="w").pack(fill="x", padx=16, pady=(10, 4))
    text_wrap = tk.Frame(panel, bg=INK_DEEP)
    text_wrap.pack(fill="both", expand=True, padx=16, pady=(0, 14))
    box = tk.Text(text_wrap, height=11, state="disabled", wrap="word",
                  font=FONT_LOG, bg=INK_DEEP, fg="#CBD8E6", relief="flat",
                  highlightthickness=0, bd=0, spacing1=1, spacing3=1)
    scroll = tk.Scrollbar(text_wrap, command=box.yview, relief="flat", bd=0,
                          width=10, troughcolor=INK_DEEP, bg="#2B4258",
                          activebackground="#3A5570", highlightthickness=0)
    box.configure(yscrollcommand=scroll.set)
    box.pack(side="left", fill="both", expand=True)
    scroll.pack(side="right", fill="y")
    box.tag_configure("good", foreground="#6FD69B")
    box.tag_configure("bad", foreground="#F0857C")

    # ---------------- logging ----------------
    # The bot runs on its own thread and tkinter must only be touched
    # from the main one, so lines are posted into a queue that the
    # window empties a few times a second.
    inbox = queue.Queue()
    counts = {"saved": 0, "issue": 0}
    last_log_time = [None]

    def log(msg):
        inbox.put(msg)

    def render(msg):
        now = time.time()
        stamp = time.strftime("%H:%M:%S")
        if last_log_time[0] is None:
            prefix = f"[{stamp}] "
        else:
            gap = now - last_log_time[0]
            prefix = f"[{stamp} +{gap:5.1f}s] "
        last_log_time[0] = now

        low = msg.lower()
        tag = None
        if "saved ->" in low:
            tag = "good"
            counts["saved"] += 1
            stat_saved.value.set(str(counts["saved"]))
        elif "error" in low or "fatal" in low:
            tag = "bad"
            counts["issue"] += 1
            stat_issue.value.set(str(counts["issue"]))

        box.configure(state="normal")
        box.insert("end", prefix + msg + "\n", tag or ())
        box.see("end")
        box.configure(state="disabled")

        stripped = msg.strip()
        if stripped and not stripped.startswith("["):
            status.configure(text=stripped[:110])

    def drain():
        try:
            while True:
                render(inbox.get_nowait())
        except queue.Empty:
            pass
        root.after(120, drain)

    def show_history():
        win = tk.Toplevel(root)
        win.title("Download history")
        win.geometry("980x480")
        win.configure(bg=PAPER)
        cols = ("when", "client", "login", "period", "types", "files", "status")
        hv = ttk.Treeview(win, columns=cols, show="headings",
                          style="Clients.Treeview")
        for c, text, w in (("when", "Date & time", 130), ("client", "Client", 200),
                           ("login", "Login ID", 120), ("period", "Period", 170),
                           ("types", "Return types", 110), ("files", "Files", 60),
                           ("status", "Status", 120)):
            hv.heading(c, text=text)
            hv.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(win, command=hv.yview)
        hv.configure(yscrollcommand=sb.set)
        hv.pack(side="left", fill="both", expand=True, padx=(16, 0), pady=16)
        sb.pack(side="left", fill="y", pady=16)
        rows = load_history()
        for h in reversed(rows):
            hv.insert("", "end", values=(
                h.get("when", ""), h.get("client", ""), h.get("login", ""),
                f"{h.get('from', '')} - {h.get('to', '')}",
                h.get("types", ""), h.get("files", 0), h.get("status", "")),
                tags=(h.get("folder", ""),))
        tk.Label(win, text="Double-click a row to open its folder.",
                 font=FONT_SMALL, fg=MUTED, bg=PAPER).place(x=16, y=456)

        def open_folder(event):
            iid = hv.identify_row(event.y)
            if not iid:
                return
            folder = hv.item(iid, "tags")[0] if hv.item(iid, "tags") else ""
            if folder and os.path.isdir(folder):
                try:
                    os.startfile(folder)
                except Exception:
                    subprocess.Popen(["xdg-open", folder])
        hv.bind("<Double-1>", open_folder)

    def start():
        d_from, d_to = e_from.get().strip(), e_to.get().strip()
        if not (d_from and d_to):
            messagebox.showwarning("Missing info",
                                   "Please fill in both date boxes.")
            return
        run_list = [c for c in clients if c["login"] in ticked_logins]
        if not run_list:
            messagebox.showwarning("No clients",
                                   "Please tick at least one client "
                                   "(use Add if the list is empty).")
            return
        ticked = [c for c, v in cat_vars.items() if v.get()]
        if not ticked:
            messagebox.showwarning("No categories",
                                   "Please tick at least one return type.")
            return
        # If every box is ticked, behave as ALL (also catches any brand-new
        # types the portal adds later)
        wanted = None if len(ticked) == len(ALL_CATEGORIES) else ticked
        types_text = "All" if wanted is None else f"{len(ticked)} selected"
        patient = bool(patient_var.get())
        background = bool(background_var.get())

        counts["saved"] = counts["issue"] = 0
        stat_saved.value.set("0")
        stat_issue.value.set("0")
        btn.configure(state="disabled")
        bar.pulse_on()
        status.configure(text="Starting Chrome...")

        def worker():
            try:
                total = len(run_list)
                for i, c in enumerate(run_list, 1):
                    log(f"\n################ CLIENT {i} of {total}: "
                        f"{c['name']} ({c['login']}) ################")
                    result = {}
                    started = datetime.datetime.now()
                    try:
                        run_bot(c["login"], c["password"], d_from, d_to,
                                wanted, log, patient=patient,
                                background=background, result=result)
                    except Exception as err:
                        log(f"UNEXPECTED ERROR: {type(err).__name__}: {err}")
                        log("Please copy this whole log to Claude.")
                    files = sum(n for _, n in result.get("summary", []))
                    add_history({
                        "when": started.strftime("%d/%m/%Y %H:%M"),
                        "client": c["name"], "login": c["login"],
                        "from": d_from, "to": d_to, "types": types_text,
                        "files": files,
                        "status": "Completed" if result.get("ok") else "Stopped early",
                        "folder": result.get("folder", ""),
                        "categories": [n for n, _ in result.get("summary", [])],
                    })
            finally:
                btn.configure(state="normal")
                bar.pulse_off()
                status.configure(
                    text=f"Finished — {counts['saved']} saved across "
                         f"{len(run_list)} client(s). Press fetch again to rerun.")
                log("\n--- Ready for another run. ---")

        threading.Thread(target=worker, daemon=True).start()

    log("Add your clients (once), tick the ones you need, set the period "
        "and press fetch returns.")
    log("Chrome will open for each client; type the security code there "
        "and click Log In.")
    log("Keep this window open to watch progress.")

    drain()
    root.mainloop()


if __name__ == "__main__":
    main()
