"""engines/wht_crosscheck.py - WHT certificates (Payee) checked against the Income Tax ledger (10.13, 10.15).

User, 8 Oct 2026 (from the ura-returns-tax-statement-download skill): 'note every withholding tax credit row of
Income Tax Level 3-2 ... flag any certificate that is not in Level 3-2'.

Inputs, both already on disk after a run:
* the certificates found for each month (Certificates step, Payee);
* the Income Tax Level 3-2 lines the Ledgers step saved ('Income Tax' and 'also download their Level 3' ticked):
  since 10.15 in Ledgers\\<TIN>\\Income Tax.xlsx (sheet 'Level 3-2', one row per line, a Period column) and,
  until folded in, Ledgers\\<TIN>\\_months\\Income Tax\\Income Tax_L3_<year>.xlsx; files of earlier versions
  (Ledgers\\<TIN>\\Income Tax_L3_*.xlsx) are read too.

The live layout (user's capture of 8 Oct 2026, Income Tax Level 3-2 for 2026-27): Payment Basis 'WHT Credit',
Payment Registration Number 'MTO...' - the withholding agent's RETURN number, not a TCC number - and Amount;
'ASYCUDA Credit' lines are customs credits (not WHT certificates) and are listed apart, for information.

Matching, in this order: the certificate's Return/Assessment Number equals the line's reference; its TCC No.
appears in the line; its Amount Withheld equals the line's amount.  A line split over two periods (same
reference) counts once.

Output: a sheet 'Check vs ledger' added to the WHT workbook - every certificate with 'Found in Income Tax
ledger' (Yes - return number / Yes - TCC No. / Yes - amount / No), the WHT credit lines no certificate matched,
then the other credit lines (information).
"""
import re
from pathlib import Path

SHEET = "Check vs ledger"
TAX = "Income Tax"
_CREDIT = re.compile(r"(?i)\bwht\b|withhold|credit|\btcc\b")
_WHT = re.compile(r"(?i)\bwht\b|withhold")


def _text(v):
    return " ".join(str(v).split()) if v is not None else ""


def _amount(v):
    from engines.wht_outputs import as_number
    try:
        n = as_number(v)
    except Exception:
        return None
    return round(float(n), 2) if n is not None else None


def ledger_files(ledgers_folder):
    """Every source of Income Tax Level 3-2 lines on disk (the workbook first)."""
    from engines import ledger_taxhead
    folder = Path(ledgers_folder)
    if not folder.exists():
        return []
    out = []
    book = folder / f"{TAX}.xlsx"
    if book.exists() and ledger_taxhead.months_in_workbook(folder, TAX):
        out.append(book)
    month_dir = ledger_taxhead.months_folder(folder, TAX)
    out += sorted(p for p in month_dir.glob(f"{TAX}_L3_*.xlsx") if not p.name.startswith("~"))
    out += sorted(p for p in folder.glob(f"{TAX}_L3_*.xlsx") if not p.name.startswith("~"))
    return out


def _find(header, *names):
    low = [_text(h).lower() for h in header]
    for n in names:
        for i, h in enumerate(low):
            if n in h:
                return i
    return None


def _structured(name, period, header, rows):
    """Lines of one period from a Level 3-2 table with a 'Payment Basis' column, one per reference."""
    basis, ref = _find(header, "payment basis"), _find(header, "registration number", "prn", "reference")
    amount = next((i for i, h in enumerate(header) if _text(h).lower() == "amount"), None)
    out, seen = [], set()
    for n, r in enumerate(rows):
        get = (lambda i: _text(r[i]) if i is not None and i < len(r) else "")
        kind = get(basis)
        if kind.lower() == "payment (prn)" or not kind:
            continue
        key = (kind.lower(), get(ref) or f"line {n}")
        if key in seen:
            continue
        seen.add(key)
        amt = _amount(get(amount))
        cells = [_text(v) for v in r if _text(v)]
        out.append({"file": name, "period": period, "kind": "wht" if _WHT.search(kind) else "other",
                    "basis": kind, "ref": get(ref), "text": " | ".join(cells),
                    "amounts": {amt} if amt else {a for a in (_amount(v) for v in cells) if a}})
    return out


def _rows(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    try:
        return [list(r) for ws in wb.worksheets[:1] for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def _workbook_lines(path, skip):
    """Lines of Income Tax.xlsx, sheet 'Level 3-2' (header: Period, Tax Position as at, the portal's columns)."""
    from openpyxl import load_workbook
    from engines import ledger_taxhead
    wb = load_workbook(path, read_only=True)
    try:
        if ledger_taxhead.SHEET not in wb.sheetnames:
            return []
        rows = [list(r) for r in wb[ledger_taxhead.SHEET].iter_rows(values_only=True)]
    finally:
        wb.close()
    if not rows:
        return []
    header = [_text(v) for v in rows[0]]
    by_period = {}
    for r in rows[1:]:
        period = ledger_taxhead.canonical(r[0]) or _text(r[0])
        if period and period not in skip:
            by_period.setdefault(period, []).append(r[2:])
    out = []
    for period, lines in by_period.items():
        out += _structured(f"{path.name} ({period})", period, header[2:], lines)
    return out


def _legacy_lines(path):
    """10.13 reading of a file whose columns are unknown: lines that mention WHT / credit / TCC."""
    rows = [[_text(v) for v in r] for r in _rows(path)]
    credit_file = "credit" in path.name.lower()
    head_at = next((i for i, r in enumerate(rows)
                    if sum(1 for v in r if v and not _amount(v)) >= 3 and i + 1 < len(rows)), None)
    out = []
    for r in rows[(head_at + 1) if head_at is not None else 0:]:
        cells = [v for v in r if v]
        if len(cells) < 2:
            continue
        text = " | ".join(cells)
        if text.lower().startswith(("tax position as at", "period:", "total")):
            continue
        if not credit_file and not _CREDIT.search(text):
            continue
        amounts = {a for a in (_amount(v) for v in cells) if a}
        if amounts:
            out.append({"file": path.name, "period": "", "kind": "wht", "basis": "", "ref": "",
                        "text": text, "amounts": amounts})
    return out


def ledger_lines(ledgers_folder):
    """[{file, period, kind ('wht'/'other'), basis, ref, text, amounts}] - every credit line of the Income Tax
    Level 3-2 on disk.  A period read from a not-yet-folded file replaces the workbook's rows of that period."""
    from engines import ledger_taxhead
    files = ledger_files(ledgers_folder)
    singles, out, from_files = [p for p in files if p.name != f"{TAX}.xlsx"], [], set()
    for path in singles:
        try:
            period, _, header, data = ledger_taxhead.read_level3(path)
            if _find(header, "payment basis") is None:
                raise ValueError("no Payment Basis column")
            out += _structured(path.name, period, header, [list(r) for r in data])
            from_files.add(period)
        except Exception:
            try:
                out += _legacy_lines(path)
            except Exception:
                continue
    book = next((p for p in files if p.name == f"{TAX}.xlsx"), None)
    if book is not None:
        try:
            out = _workbook_lines(book, from_files) + out
        except Exception:
            pass
    return out


def ledger_credit_lines(ledgers_folder):
    """The WHT credit lines only (10.13 name, kept for callers)."""
    return [ln for ln in ledger_lines(ledgers_folder) if ln["kind"] == "wht"]


def _col(df, *words):
    for c in df.columns:
        if all(w in str(c).lower() for w in words):
            return c
    return None


def crosscheck(months, ledgers_folder):
    """(certificate table with a 'Found in Income Tax ledger' column, unmatched WHT credit lines, files used,
    other credit lines) or None when no Income Tax Level 3-2 is on disk."""
    import pandas as pd
    files = ledger_files(ledgers_folder)
    if not files:
        return None
    every = ledger_lines(ledgers_folder)
    lines = [ln for ln in every if ln["kind"] == "wht"]
    others = [ln for ln in every if ln["kind"] != "wht"]
    frames = []
    for r in months:
        if r.df is not None and len(r.df):
            d = r.df.copy()
            d.insert(0, "Month", r.title)
            frames.append(d)
    certs = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    used = set()
    found = []
    tcc_col, amt_col = (_col(certs, "tcc"), _col(certs, "withheld")) if len(certs) else (None, None)
    ret_col = _col(certs, "return") if len(certs) else None
    for _, c in certs.iterrows():
        hit, how = None, ""
        ret = _text(c.get(ret_col)) if ret_col else ""
        if ret:
            hit = next((i for i, ln in enumerate(lines) if i not in used and
                        (ln["ref"] == ret or (not ln["ref"] and ret in ln["text"]))), None)
            how = "Yes - return number"
        tcc = _text(c.get(tcc_col)) if tcc_col else ""
        if hit is None and tcc:
            hit = next((i for i, ln in enumerate(lines) if i not in used and tcc in ln["text"]), None)
            how = "Yes - TCC No."
        if hit is None:
            amt = _amount(c.get(amt_col)) if amt_col else None
            if amt:
                hit = next((i for i, ln in enumerate(lines) if i not in used and amt in ln["amounts"]), None)
                how = "Yes - amount"
        if hit is None:
            found.append(("No", ""))
        else:
            used.add(hit)
            found.append((how, lines[hit]["file"]))
    if len(certs):
        certs["Found in Income Tax ledger"] = [f for f, _ in found]
        certs["Ledger file"] = [f for _, f in found]
    left = [ln for i, ln in enumerate(lines) if i not in used]
    return certs, left, [p.name for p in files], others


def add_sheet(workbook_path, result):
    """Add (or replace) the 'Check vs ledger' sheet in the WHT workbook."""
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    wb = load_workbook(workbook_path)
    if SHEET in wb.sheetnames:
        del wb[SHEET]
    ws = wb.create_sheet(SHEET, 1)
    bold = Font(bold=True)

    def put(row, values, strong=False):
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=row, column=c, value=v)
            if isinstance(v, str):
                cell.data_type = "s"
            if strong:
                cell.font = bold
    if result is None:
        put(1, ["Not checked: no Income Tax Level 3 file in the client's Ledgers folder. Tick 'Income Tax' and "
                "'also download their Level 3 transactions' on the Ledgers page, then run again."], True)
        wb.save(workbook_path)
        return
    certs, left, files = result[:3]
    others = result[3] if len(result) > 3 else []
    n_no = int((certs["Found in Income Tax ledger"] == "No").sum()) if len(certs) else 0
    put(1, [f"{len(certs)} certificate(s); {len(certs) - n_no} found in the Income Tax ledger, {n_no} not found; "
            f"{len(left)} ledger WHT credit line(s) without a certificate."], True)
    put(2, ["Ledger files read: " + ", ".join(files)])
    r = 4
    if len(certs):
        put(r, [str(c) for c in certs.columns], True)
        for vals in certs.itertuples(index=False):
            r += 1
            put(r, [("" if v is None else v) for v in vals])
        r += 2
    put(r, ["Income Tax ledger WHT credit lines with no certificate"], True)
    put(r + 1, ["Ledger file", "Line"], True)
    for i, ln in enumerate(left, r + 2):
        put(i, [ln["file"], ln["text"]])
    if others:
        r = r + 3 + len(left)
        put(r, ["Other credits in the Income Tax ledger (not WHT certificates - for information)"], True)
        put(r + 1, ["Ledger file", "Line"], True)
        for i, ln in enumerate(others, r + 2):
            put(i, [ln["file"], ln["text"]])
    for col in "ABCDEFGHIJKLMN":
        ws.column_dimensions[col].width = 18
    ws.freeze_panes = "A5"
    wb.save(workbook_path)
