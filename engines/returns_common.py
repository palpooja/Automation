"""
engines/returns_common.py - the Selenium-free parts of the Returns
Downloader v4.3 (ura_returns_bot.pyw), extracted unchanged: return-type
lists, file naming, and the Returns_Status.xlsx writer.  Used by
engines/returns_pw.py and modules/returns.py.  returns_engine.py (the old
Selenium program) is no longer imported by the suite.
"""
import os
import time

try:                              # for the Returns_Status.xlsx sheet
    import openpyxl
    from openpyxl.styles import Font
except ImportError:               # pragma: no cover
    openpyxl = None

BOT_TYPE, BOT_DL, BOT_CHK = "Return type", "Downloaded", "Checked on"
BOT_FILE = "Saved as"          # 10.13: the file name, now that names no longer carry the return number

TAB_WAIT_SECONDS = 60              # normal mode: for the return's tab to open

PAGE_WAIT_SECONDS = 60             # normal mode: for the Print button to appear

PATIENT_LIMIT_SECONDS = 15 * 60    # "wait until it downloads" safety valve
# 10.14 (user, 8 Oct 2026 - a VAT return never opened, the run waited until Chrome was closed): on the first
# try a return gets at most FIRST_TRY_SECONDS from View to its Print button; a slower one is set aside and tried
# again with the full wait after the client's other steps.  Live: the slowest return that DID open took 64 s
# to open (6 Oct) and 70 s to finish loading (7 Oct); normal ones 1-3 s.
FIRST_TRY_SECONDS = 120
SKIP_OFFER_AFTER_SECONDS = 60      # second try: 'Attention needed' + Skip / Keep waiting after this long

HTML_CATEGORIES = {"Local Excise Duty Return"}   # also saved as .mhtml

STATUS_XLSX = "Returns_Status.xlsx"            # one per client, all return types

BRANCH_CODES = {
    "LEDRT",     # Local Excise Duty Return
    "VATRT",     # Monthly VAT Return
    "PAYERT",    # PAYE Return
    "WTHDRT",    # Withholding Tax Return
}

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

DEFAULT_TICKED = {
    "Monthly VAT Return",
    "PAYE Return",
    "Income Tax Return for Individual",
    "Income Tax Return for Non-Individual",
    "Income Tax Provisional Return for Individual",
    "Provisional Return for Non-Individual",
    "Local Excise Duty Return",
    "Withholding Tax Return",       # 8 Oct 2026: was left unticked (default carried over from 4.3)
}

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

def _clean(text):
    return " ".join((text or "").split())

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
                    if h and h not in (BOT_TYPE, BOT_DL, BOT_CHK, BOT_FILE)]

    # ---- this run's rows
    new_rows = []
    for e in entries:
        d = {BOT_TYPE: e["type"]}
        for h, v in zip(e["heads"], e["cells"]):
            d[h] = v
            if h not in portal_heads:
                portal_heads.append(h)
        d[BOT_DL] = "Yes" if e["downloaded"] else "No"
        d[BOT_FILE] = e.get("saved_as", "")
        d[BOT_CHK] = checked_on
        new_rows.append(d)
    heads = [BOT_TYPE] + portal_heads + [BOT_DL, BOT_FILE, BOT_CHK]

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
    from engines.sheet_safety import neutralise_sheet
    for sheet in (ws, ws2):                 # portal text stays text (finding L2)
        neutralise_sheet(sheet)
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


# ---------------------------------------------------------------- friendly file names (10.13)
# User, 8 Oct 2026 (from the ura-returns-tax-statement-download skill): name returns for the audit file -
# '<n>. <Month> <YYYY>.pdf' with n = the month's place in the period (first month = 1), the month taken from
# the Return Period From date; an amended return '... - Amended filed <dd-mm-yyyy>.pdf'; a branch other than
# the main one ' - <Branch>'; yearly returns 'Original Prov 2025-26.pdf' / 'Amended Prov 2025-26 filed ...'.
# The return number is kept in _saved_as.json (name -> return number) and in Returns_Status.xlsx
# ('Saved as'), so 'already saved' never depends on the name.
import calendar as _cal
import datetime as _dt
import json as _json
import re as _re

SAVED_AS_FILE = "_saved_as.json"
MONTH_NAMES = list(_cal.month_name)
_DATE = _re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")
_RETNO = _re.compile(r"^[A-Z0-9]{10,}$")


def _d(text):
    m = _DATE.match(text or "")
    if not m:
        return None
    try:
        return _dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def fy_label(date_from, date_to):
    """'01/07/2025', '30/06/2026' -> '2025-26'; any other span -> '01-06-2025 to 31-07-2026'."""
    a, b = _d(date_from), _d(date_to)
    if a and b and (a.month, a.day) == (7, 1) and (b.month, b.day) == (6, 30) and b.year == a.year + 1:
        return f"{a.year}-{str(b.year)[2:]}"
    return f"{(date_from or '').replace('/', '-')} to {(date_to or '').replace('/', '-')}"


def return_facts(cells, heads=None):
    """What the return-history row says: return number, period from/to, applied date, Original/Amended, branch."""
    norm = [" ".join(str(c).split()) for c in cells]
    heads = [str(h or "").lower() for h in (heads or [])]
    retno = next((c for c in norm if _RETNO.match(c) and _re.search(r"[A-Z]", c) and _re.search(r"\d", c)), "")
    dates = [c for c in norm if _DATE.match(c)]
    kind = next((c.split()[0].capitalize() for c in norm
                 if _re.match(r"(?i)^(original|amended)\b", c)), "")
    branch = ""
    for i, h in enumerate(heads):
        if "branch" in h and i < len(norm) and norm[i]:
            branch = norm[i]
            break
    else:
        branch = next((c for c in norm if _re.search(r"(?i)\bbranch\b", c)), "")
    return {"retno": retno, "from": dates[0] if dates else "", "to": dates[1] if len(dates) > 1 else "",
            "applied": dates[2] if len(dates) > 2 else "", "kind": kind, "branch": branch}


def friendly_label(cells, heads, return_type, period_from):
    """The file name (no extension) for one return, or '' when the row does not give a period."""
    f = return_facts(cells, heads)
    a, b, start = _d(f["from"]), _d(f["to"]), _d(period_from)
    if not (a and b):
        return ""
    amended = f["kind"] == "Amended"
    filed = f" filed {f['applied'].replace('/', '-')}" if f["applied"] else ""
    branch = ""
    if f["branch"] and f["branch"].strip().upper() not in ("MAIN BRANCH", "MAIN", "HEADQUARTER", "HQ"):
        branch = f" - {safe_name(f['branch'])}"
    monthly = (a.day == 1 and a.year == b.year and a.month == b.month
               and b.day == _cal.monthrange(b.year, b.month)[1])
    if monthly:
        n = (a.year - start.year) * 12 + (a.month - start.month) + 1 if start else a.month
        base = f"{n}. {MONTH_NAMES[a.month]} {a.year}"
        return safe_name(base + (f" - Amended{filed}" if amended else "") + branch)
    short = ("Prov" if "provisional" in return_type.lower()
             else "ITR" if "income tax return" in return_type.lower() else "Return")
    span = fy_label(f["from"], f["to"])
    return safe_name(f"{'Amended' if amended else 'Original'} {short} {span}" + (filed if amended else "") + branch)


def load_saved_as(folder):
    try:
        with open(os.path.join(folder, SAVED_AS_FILE), encoding="utf-8") as fh:
            data = _json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_saved_as(folder, mapping):
    path = os.path.join(folder, SAVED_AS_FILE)
    tmp = path + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as fh:
            _json.dump(mapping, fh, indent=1, sort_keys=True)
        os.replace(tmp, path)
    except OSError:
        pass


def unique_label(label, retno, mapping, applied=""):
    """`label`, unless another return already owns that name; then ' (filed dd-mm-yyyy)' or ' (2)', ' (3)'..."""
    owner = mapping.get(label)
    if not owner or owner == retno:
        return label
    if applied:
        alt = f"{label} (filed {applied.replace('/', '-')})"
        if mapping.get(alt) in (None, retno):
            return alt
    n = 2
    while mapping.get(f"{label} ({n})") not in (None, retno):
        n += 1
    return f"{label} ({n})"
