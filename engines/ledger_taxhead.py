"""engines/ledger_taxhead.py - one workbook per tax head, updated in place.

User, 6 Oct 2026: 'one file per Tax Head. Make only one table and keep updating data instead of making a
new file every time. Also make sure it is not taking extra time.'
User, 8 Oct 2026: 'I don't want an Excel for every month, I want a combined one like WHT Receivables.'

    Ledgers\\<TIN>\\<Tax head>.xlsx
        Summary      one row per month: Period, Status, Payment (Level 2), Level 3 rows, Level 3 amount,
                     Tax Position as at, Last checked, Detail  (an Excel table)
        Level 3-2    ALL months in ONE Excel table: Period | Tax Position as at | the portal's Level 3-2
                     columns (Sr No., Payment Basis, ...) - unchanged from 10.9
        JUL-2025 ... one tab per month, oldest first: the portal's Level 3 exactly as exported
                     (same cells, merged headings, number formats)
    - the same layout as the WHT workbook (Summary, Combined, a tab per month).

The portal's monthly downloads are still saved first to Ledgers\\<TIN>\\_months\\<Tax head>\\ during the
step (resume after a stop, the background-route proof, the portal-file conversion check).  At the end of
the step they are folded into the workbook and REMOVED - each one only after the workbook has been
written and read back with that month's tab in it.  If the workbook cannot be written (open in Excel),
nothing is removed and the next run folds them in.  KEEP_MONTH_FILES / keep_month_files=True keeps them.

Updating: a month with a monthly file takes its rows and tab from that file (the newest download wins);
other months keep their rows and tab from the workbook.  Values are copied as the converter stored them
(amounts as numbers, PRNs and dates as text; text is never a formula).

10.15 - yearly tax heads (Income Tax, Rental) use the same workbook: a "period" is a month (MAY-2026) or a
year as the portal writes it (JUL-2025 to JUN-2026), one Summary row and one tab per year.

10.15 - the check (user's live output of 8 Oct 2026, every VAT, LED, PAYE and IT-WHT month read line by line):
  Level 2's "Payment" heading has sub-columns (VAT: PRN | WHT Credit | Offset Received(Other Period);
  Income Tax: PRN | Tax Credit; IT-WHT, PAYE, LED: one column).  Their TOTAL equals the Amount of every
  DIFFERENT reference on the period's Level 3-2 lines - a PRN split over two periods is listed twice with the
  same Amount and counts once.  Seen live: VAT MAY-2026 (one PRN + one Credit Balance
  Transfer); VAT FEB-2026 two PRNs + a transfer; LED FEB-2026 one PRN listed twice; Income Tax 2026-27
  PRN + Tax Credit (seven PRNs; a WHT Credit and an ASYCUDA Credit).  The split between the sub-columns
  is NOT always the Level 3 split: VAT AUG-2025 and SEP-2025 show an old Credit Balance Transfer (effective
  01-JUL-2021) inside the PRN sub-column - so only the total is checked; both splits are shown.
  The 10.13 rule (total Principal) was wrong: part of a payment can stay unallocated (Principal less than the
  Amount) and Level 2 still counts all of it.  "Not allocated" = per PRN, Amount minus Principal, Penalty,
  Interest and Refund - information only, never a check.
  The Level 2 amounts are read from the saved <Tax head>_L2.xlsx; when it cannot be read, from the amounts the
  run noted ('Payment 350,000; Tax credit 460,000').
"""
import datetime
import os
import re
from pathlib import Path

MONTHS_DIR = "_months"
SHEET = "Level 3-2"
SUMMARY = "Summary"
FIXED = ["Period", "Tax Position as at"]
SUMMARY_HEAD = ["Period", "Status", "Payment (Level 2)", "Level 3 rows", "Level 3 total", "Difference", "Check",
                "Level 2 split", "Level 3 PRN payments", "Level 3 credits", "Credits by type", "Not allocated",
                "Tax Position as at", "Last checked", "Detail"]
CHECK_TOLERANCE = 1                   # UGX: rounding only
PAYMENT_BASIS = "payment (prn)"       # the Level 3-2 'Payment Basis' of a payment (live captures)
_ALLOCATED = ("Principal", "Penalty", "Interest", "Refund")
IN_WORKBOOK = "Level 3 saved"
_MONTH = re.compile(r"^([A-Z]{3})-(\d{4})$")
_YEAR = re.compile(r"^([A-Z]{3}-\d{4})\s+TO\s+([A-Z]{3}-\d{4})$")      # JUL-2025 to JUN-2026
_ABBR = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
_HEAD_WORDS = ("sr no.", "sr no", "s/n", "payment basis")


def months_folder(save_folder, tax_type) -> Path:
    return Path(save_folder) / MONTHS_DIR / tax_type


def move_old_month_files(save_folder, tax_type) -> int:
    """Monthly files saved by earlier versions next to the workbooks move into _months, so they
    are recognised as done and folded into the workbook.  Returns how many moved."""
    save_folder, dest = Path(save_folder), months_folder(save_folder, tax_type)
    moved = 0
    for path in save_folder.glob(f"{tax_type}_L3_*.xlsx"):
        dest.mkdir(parents=True, exist_ok=True)
        if not (dest / path.name).exists():
            os.replace(path, dest / path.name)
            moved += 1
    return moved


def is_month(name) -> bool:
    m = _MONTH.match(str(name or "").strip().upper())
    return bool(m and m.group(1) in _ABBR)


def canonical(name) -> str:
    """'MAY-2026' for a month, 'JUL-2025 to JUN-2026' for a year (as the portal writes it), '' otherwise."""
    text = " ".join(str(name or "").split()).upper()
    if is_month(text):
        return text
    m = _YEAR.match(text)
    if m and is_month(m.group(1)) and is_month(m.group(2)):
        return f"{m.group(1)} to {m.group(2)}"
    return ""


def is_period(name) -> bool:
    """A month or a year period (yearly tax heads: Income Tax, Rental)."""
    return bool(canonical(name))


def _key(period):
    text = canonical(period)
    m = _MONTH.match(text.split(" ")[0]) if text else None
    return (int(m.group(2)), _ABBR.index(m.group(1)) + 1, len(text)) if m else (9999, 99, 0)


def _text(v):
    return " ".join(str(v).split()) if v is not None else ""


def _number(v):
    """An amount as a number (None when it is not one)."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    from engines.excel_files import cell_value
    n = cell_value(_text(v)) if _text(v) else None
    return n if isinstance(n, (int, float)) else None


# ---------------------------------------------------------------- reading
def _parse_level3(rows, name_hint=""):
    """(period, position time, header, data rows) from the rows of one Level 3 sheet."""
    period = when = ""
    head_at = None
    for i, r in enumerate(rows):
        first = next((_text(v) for v in r if _text(v)), "")
        if first.lower().startswith("tax position as at"):
            when = first[len("tax position as at"):].strip()
        elif first.lower().startswith("period:"):
            period = canonical(first.split(":", 1)[1]) or first.split(":", 1)[1].strip().upper()
        elif any(_text(v).lower() in _HEAD_WORDS for v in r):
            head_at = i
            break
    if head_at is None:
        raise ValueError(f"{name_hint}: no Level 3 heading row (Sr No. / Payment Basis)")
    header = [_text(v) for v in rows[head_at]]
    while header and not header[-1]:
        header.pop()
    data = []
    for r in rows[head_at + 1:]:
        r = list(r[:len(header)]) + [None] * (len(header) - len(r))
        if any(_text(v) for v in r):
            data.append(r)
    return period, when, header, data


def read_level3(path):
    """(period, position time, header, data rows) of one monthly Level 3 file as saved."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    try:
        rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()
    period, when, header, data = _parse_level3(rows, Path(path).name)
    if not period:                             # name is <tax>_L3_<MON-YYYY>.xlsx
        stem = Path(path).stem.rsplit("_L3_", 1)[-1]
        period = canonical(stem) or stem.upper()
    return period, when, header, data


def _existing(path):
    """What a workbook written earlier holds: (months, summary, workbook or None).
    months = {period: (header, [(when, values)])}; summary = {period: {column: value}}.
    The workbook is returned open (normal mode) so its month tabs can be copied."""
    from openpyxl import load_workbook
    if not Path(path).exists():
        return {}, {}, None
    wb = load_workbook(path)
    months, summary = {}, {}
    if SHEET in wb.sheetnames:
        rows = [list(r) for r in wb[SHEET].iter_rows(values_only=True)]
        if rows:
            header = [_text(v) for v in rows[0]][len(FIXED):]
            for r in rows[1:]:
                if not any(_text(v) for v in r):
                    continue
                months.setdefault(canonical(r[0]) or _text(r[0]).upper(), (header, []))[1].append(
                    (r[1], list(r[2:2 + len(header)])))
    if SUMMARY in wb.sheetnames:
        rows = [list(r) for r in wb[SUMMARY].iter_rows(values_only=True)]
        if rows:
            head = [_text(v) for v in rows[0]]
            for r in rows[1:]:
                if r and is_period(r[0]):
                    summary[canonical(r[0])] = dict(zip(head, r))
    return months, summary, wb


_DONE_CACHE = {}


def months_in_workbook(save_folder, tax_type) -> set:
    """Months whose Level 3 is inside <tax_type>.xlsx (a month tab or rows in the Level 3-2 table).
    These count as downloaded although their monthly file has been removed."""
    from openpyxl import load_workbook
    path = Path(save_folder) / f"{tax_type}.xlsx"
    try:
        stamp = path.stat().st_mtime_ns
    except OSError:
        return set()
    hit = _DONE_CACHE.get(str(path))
    if hit and hit[0] == stamp:
        return hit[1]
    found = set()
    try:
        wb = load_workbook(path, read_only=True)
        try:
            found = {canonical(n) for n in wb.sheetnames if is_period(n)}
            if SHEET in wb.sheetnames:
                for r in wb[SHEET].iter_rows(min_row=2, max_col=1, values_only=True):
                    if r and is_period(r[0]):
                        found.add(canonical(r[0]))
        finally:
            wb.close()
    except Exception:
        return set()
    _DONE_CACHE[str(path)] = (stamp, found)
    return found


# ---------------------------------------------------------------- writing
def update_tax_head(save_folder, tax_type, log=print, statuses=None, keep_month_files=False, checks=None):
    """Write/update <save_folder>/<tax_type>.xlsx.  `statuses` = this run's outcome per month
    {period: (status, detail)} for the Summary.  Returns the path, or None when there is nothing
    to write.  Monthly files folded in are removed unless keep_month_files."""
    from openpyxl import load_workbook
    from engines.excel_files import excel_ok
    save_folder = Path(save_folder)
    target = save_folder / f"{tax_type}.xlsx"
    months, summary, old_wb = _existing(target)                  # earlier runs
    tabs = {}                                                    # period -> source worksheet
    if old_wb is not None:
        for name in old_wb.sheetnames:
            if is_period(name):
                tabs[canonical(name)] = old_wb[name]
    folded, open_books = [], []
    try:
        for path in sorted(months_folder(save_folder, tax_type).glob(f"{tax_type}_L3_*.xlsx")):
            if not excel_ok(path):
                continue
            try:
                period, when, header, data = read_level3(path)
                book = load_workbook(path)
            except Exception as e:
                log(f"  {path.name}: not added to {target.name} ({e})")
                continue
            open_books.append(book)
            months[period] = (header, [(when, r) for r in data])     # this run's download wins
            tabs[period] = book.worksheets                           # every sheet of the portal's file
            folded.append((period, path))
        statuses = {canonical(k) or k: v for k, v in (statuses or {}).items()}
        if not months and not tabs and not summary and not statuses:
            return None
        rows_total = sum(len(v[1]) for v in months.values())
        columns = []                                                 # union of headings, in order
        for p in sorted(months, key=_key):
            for h in months[p][0]:
                if h and h not in columns:
                    columns.append(h)
        summary = _summary(summary, months, statuses, level2_payments(save_folder, tax_type))
        if checks is not None:                                   # months fetched or re-read this run
            for period in sorted(summary, key=_key):
                c = summary[period].get("Check")
                if c and c != "OK" and (period in statuses or any(p == period for p, _ in folded)):
                    checks.append((period, c))
        _write(target, columns, months, tabs, summary)
    finally:
        for b in open_books:
            b.close()
        if old_wb is not None:
            old_wb.close()
    log(f"Updated: {target}  ({len(months)} month(s), {rows_total} row(s) in one table, "
        f"a tab per month)")
    if folded and not keep_month_files:
        removed = remove_folded(target, folded)
        if removed:
            log(f"  {removed} monthly file(s) are now inside {target.name} - removed from {MONTHS_DIR}.")
    return target


def _detail_amounts(detail):
    """(payment, tax credit, rest of the text) from a run's detail 'Payment 350,000; Tax credit 460,000'."""
    pay = credit = None
    rest = []
    for part in str(detail or "").split(";"):
        part = part.strip()
        m = re.match(r"(?i)^payment\s+(.+)$", part)
        c = re.match(r"(?i)^tax credit\s+(.+)$", part)
        if m:
            pay = _number(m.group(1)) if _number(m.group(1)) is not None else m.group(1)
        elif c:
            credit = _number(c.group(1)) if _number(c.group(1)) is not None else c.group(1)
        elif part:
            rest.append(part)
    return pay, credit, "; ".join(rest)


def _col(header, *names):
    low = [h.lower() for h in header]
    for n in names:
        if n.lower() in low:
            return low.index(n.lower())
    for i, h in enumerate(low):
        if any(n.lower() in h for n in names):
            return i
    return None


def level3_figures(header, rows):
    """{'payments', 'prns', 'credits', 'by_type', 'not_allocated'} of one period's Level 3-2 lines.
    payments = Amount of each different PRN on the 'Payment (PRN)' lines (a PRN listed twice counts once);
    credits  = the same for every other kind of line, also per kind (by_type);
    not_allocated = per PRN, Amount minus Principal, Penalty, Interest and Refund (information only)."""
    basis = _col(header, "Payment Basis")
    ref = _col(header, "Payment Registration Number", "PRN")
    amount = _col(header, "Amount")
    alloc = [i for i in (_col(header, a) for a in _ALLOCATED) if i is not None]
    if amount is None:
        return None
    seen, spent = {}, {}
    for n, (_, v) in enumerate(rows):
        get = (lambda i: v[i] if i is not None and i < len(v) else None)
        kind = _text(get(basis)) or "Payment (PRN)"
        key = (kind.lower(), _text(get(ref)) or f"line {n}")
        amt = _number(get(amount))
        if amt is None:
            continue
        seen.setdefault(key, (kind, amt))
        spent[key] = spent.get(key, 0) + sum(_number(get(i)) or 0 for i in alloc)
    payments = sum(a for (k, _), (_, a) in seen.items() if k == PAYMENT_BASIS)
    by_type = {}
    for (k, _), (kind, a) in seen.items():
        if k != PAYMENT_BASIS:
            by_type[kind] = by_type.get(kind, 0) + a
    left = sum(max(0, a - spent[key]) for key, (_, a) in seen.items() if key[0] == PAYMENT_BASIS) if alloc else None
    return {"payments": payments, "prns": sum(1 for k, _ in seen if k == PAYMENT_BASIS),
            "credits": sum(by_type.values()), "by_type": by_type, "not_allocated": left}


def level2_payments(save_folder, tax_type):
    """{period: [(sub-column, amount), ...]} - the 'Payment' columns of the saved <tax>_L2.xlsx (two heading
    rows: 'Payment' over PRN | WHT Credit | ...).  {} when the file is missing or unreadable."""
    from openpyxl import load_workbook
    path = Path(save_folder) / f"{tax_type}_L2.xlsx"
    if not path.exists():
        return {}
    try:
        wb = load_workbook(path, read_only=True)
        try:
            rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
        finally:
            wb.close()
    except Exception:
        return {}
    h = next((i for i, r in enumerate(rows) if any(_text(v).lower() == "period" for v in r)), None)
    if h is None:
        return {}
    top = [_text(v) for v in rows[h]]
    sub = [_text(v) for v in rows[h + 1]] if h + 1 < len(rows) else []
    if sub and any(is_period(v) for v in sub):
        sub = []                                         # no second heading row
    pay = next((i for i, v in enumerate(top) if v.lower() == "payment"), None)
    per = next(i for i, v in enumerate(top) if v.lower() == "period")
    if pay is None:
        return {}
    stop = next((i for i in range(pay + 1, len(top)) if top[i]), len(top))
    cols = [i for i in range(pay, stop) if i < len(sub) and sub[i]] or [pay]
    names = [sub[i] if i < len(sub) and sub[i] else "Payment" for i in cols]
    out = {}
    for r in rows[h + 1:]:
        if per < len(r) and is_period(r[per]):
            out[canonical(r[per])] = [(n, _number(r[i]) if i < len(r) else None) for n, i in zip(names, cols)]
    return out


def _summary(old, months, statuses, level2=None):
    """{period: row dict} - earlier rows kept, this run's outcome and the workbook's contents win."""
    today = datetime.date.today().strftime("%d/%m/%Y")
    out = {p: dict(r) for p, r in old.items()}
    level2 = level2 or {}
    for period, (status, detail) in statuses.items():
        if not is_period(period):
            continue
        r = out.setdefault(canonical(period), {})
        r["Status"] = status
        pay, credit, rest = _detail_amounts(detail)
        noted = [(n, v) for n, v in (("PRN", pay), ("Tax Credit", credit)) if v is not None]
        if noted and canonical(period) not in level2:          # the saved Level 2 file wins
            _level2_into(r, noted)
            r["_noted"] = True
        r["Detail"] = rest or None
        r["Last checked"] = today
    for period, parts in level2.items():
        if period in out or period in months:
            _level2_into(out.setdefault(period, {}), parts)
    for period, (header, rows) in months.items():
        r = out.setdefault(period, {})
        r["Status"] = IN_WORKBOOK                   # it IS in the workbook (this run or an earlier one)
        r["Level 3 rows"] = len(rows)
        fig = level3_figures(header, rows)
        if fig is not None:
            r["Level 3 total"] = fig["payments"] + fig["credits"]
            r["Level 3 PRN payments"] = fig["payments"]
            r["Level 3 credits"] = fig["credits"] or None
            r["Credits by type"] = "; ".join(f"{k} {v:,.0f}" for k, v in fig["by_type"].items()) or None
            r["Not allocated"] = fig["not_allocated"] or None
        r["Tax Position as at"] = next((w for w, _ in rows if w), r.get("Tax Position as at"))
        r.setdefault("Last checked", today)
        r["Difference"], r["Check"] = _check(period, r, header, rows, fig)
    return out


def _level2_into(r, parts):
    nums = [(n, v) for n, v in parts if isinstance(v, (int, float))]
    if not nums:
        if parts:
            r["Payment (Level 2)"] = parts[0][1]
        return
    r["Payment (Level 2)"] = sum(v for _, v in nums)
    r["Level 2 split"] = "; ".join(f"{n} {v:,.0f}" for n, v in nums if v) if len(nums) > 1 else None


def _check(period, r, header, rows, fig=None):
    """(difference, 'OK' / 'check: ...') - the Level 2 Payment total of the period against its Level 3-2
    lines (rule in the module text)."""
    pay = r.get("Payment (Level 2)")
    if fig is None or not isinstance(pay, (int, float)) or not rows:
        return None, None
    total = fig["payments"] + fig["credits"]
    diff = pay - total
    if abs(diff) < CHECK_TOLERANCE:
        return 0, "OK"
    if r.get("_noted") and abs(pay - fig["payments"]) < CHECK_TOLERANCE:
        return 0, "OK"            # noted from the screen's first Payment column only (PRN)
    why = (f"check: Level 2 payment {pay:,.0f}, Level 3 lines {total:,.0f} (PRN payments "
           f"{fig['payments']:,.0f}, {fig['prns']} PRN(s)"
           + (f"; credits {fig['credits']:,.0f}" if fig["credits"] else "") + f"), differs by {diff:,.0f}")
    adj = _col(header, "Adjusted Against Period")
    if adj is not None:
        mine = canonical(period)
        other = sorted({_text(v[adj]).upper() for _, v in rows
                        if adj < len(v) and _text(v[adj]) and _text(v[adj]).upper() not in (mine.upper(), "NA")})
        if other:
            why += f" - line(s) adjusted against {', '.join(other)}"
    return diff, why


def _safe_name(text):
    return re.sub(r"[^A-Za-z0-9_]", "_", text)[:250]


def _cell(ws, row, col, value):
    """Write one value; text always stays text (never a formula)."""
    cell = ws.cell(row=row, column=col, value=value)
    if isinstance(value, str):
        cell.data_type = "s"
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        cell.number_format = "#,##0" if isinstance(value, int) else "#,##0.00"
    return cell


def _table(ws, head, name):
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
    for c in range(1, len(head) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)
    ws.freeze_panes = "A2"
    table = Table(displayName=_safe_name(name), ref=f"A1:{get_column_letter(len(head))}{max(2, ws.max_row)}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)


def _widths(ws, width):
    from openpyxl.utils import get_column_letter
    for c, w in width.items():
        ws.column_dimensions[get_column_letter(c)].width = min(45, max(9, w + 2))


def _copy_sheets(sources, dst):
    """Copy one month's portal sheet(s) into dst: values, bold, number formats, merges, widths.
    A file with several sheets has them one under the other, a blank row between."""
    from openpyxl.cell.cell import MergedCell
    from openpyxl.styles import Font
    if not isinstance(sources, (list, tuple)):
        sources = [sources]
    offset, head_row = 0, None
    for src in sources:
        for row in src.iter_rows():
            for c in row:
                if isinstance(c, MergedCell) or c.value is None:
                    continue
                d = _cell(dst, c.row + offset, c.column, c.value)
                if c.number_format and c.number_format != "General":
                    d.number_format = c.number_format
                if c.font is not None and c.font.b:
                    d.font = Font(bold=True)
                if head_row is None and _text(c.value).lower() in _HEAD_WORDS:
                    head_row = c.row + offset
        for rng in src.merged_cells.ranges:
            dst.merge_cells(start_row=rng.min_row + offset, start_column=rng.min_col,
                            end_row=rng.max_row + offset, end_column=rng.max_col)
        for key, dim in src.column_dimensions.items():
            if dim.width:
                dst.column_dimensions[key].width = max(dst.column_dimensions[key].width or 0, dim.width)
        offset += src.max_row + 1
    if head_row:
        dst.freeze_panes = f"A{head_row + 1}"


def _rebuilt_tab(ws, period, header, rows):
    """A month known only from the Level 3-2 table (its file gone): heading lines, header, rows."""
    from openpyxl.styles import Font
    when = next((w for w, _ in rows if w), "")
    _cell(ws, 1, 1, f"Tax Position as at {when}".strip())
    _cell(ws, 2, 1, f"Period: {period}")
    for c, h in enumerate(header, 1):
        _cell(ws, 3, c, h).font = Font(bold=True)
    for i, (_, values) in enumerate(rows, 4):
        for c, v in enumerate(values, 1):
            if v is not None:
                _cell(ws, i, c, v)
    _cell(ws, len(rows) + 5, 1, "Rebuilt from the Level 3-2 table (the portal's monthly file was not on disk).")
    ws.freeze_panes = "A4"


def _write(target, columns, months, tabs, summary):
    from openpyxl import Workbook
    wb = Workbook()
    # ---- Summary
    ws = wb.active
    ws.title = SUMMARY
    ws.append(SUMMARY_HEAD)
    width = {i: len(h) for i, h in enumerate(SUMMARY_HEAD, 1)}
    for r, period in enumerate(sorted(summary, key=_key), 2):
        row = summary[period]
        for c, h in enumerate(SUMMARY_HEAD, 1):
            v = period if c == 1 else row.get(h)
            if v is not None:
                _cell(ws, r, c, v)
                width[c] = max(width[c], len(f"{v:,}" if isinstance(v, (int, float)) else str(v)))
    _widths(ws, width)
    _table(ws, SUMMARY_HEAD, f"Summary_{target.stem}")
    # ---- Level 3-2: every month in one table (unchanged from 10.9)
    ws = wb.create_sheet(SHEET)
    head = FIXED + columns
    ws.append(head)
    width = {i: len(h) for i, h in enumerate(head, 1)}
    for period in sorted(months, key=_key):
        header, rows = months[period]
        pos = {h: i for i, h in enumerate(header)}
        for when, values in rows:
            line = [period, when] + [values[pos[h]] if h in pos and pos[h] < len(values) else None
                                     for h in columns]
            r = ws.max_row + 1
            for c, v in enumerate(line, 1):
                _cell(ws, r, c, v)
                width[c] = max(width[c], len(str(v)) if v is not None else 0)
    _widths(ws, width)
    _table(ws, head, f"L3_{target.stem}")
    # ---- one tab per month, oldest first
    for period in sorted(set(tabs) | set(months), key=_key):
        ws = wb.create_sheet(period)
        if period in tabs:
            _copy_sheets(tabs[period], ws)
        else:
            _rebuilt_tab(ws, period, *months[period])
    tmp = target.with_name("~" + target.stem + ".tmp.xlsx")
    target.parent.mkdir(parents=True, exist_ok=True)
    wb.save(tmp)
    try:
        os.replace(tmp, target)                                  # one file, replaced in one step
    except PermissionError:                                      # open in Excel: keep the new data
        alt = target.with_name(f"{target.stem} (updated {datetime.datetime.now():%Y%m%d_%H%M}).xlsx")
        os.replace(tmp, alt)
        raise PermissionError(f"{target.name} is open in Excel - this run's update was saved as {alt.name}; "
                              "close the file and run again to update it")


def remove_folded(target, folded) -> int:
    """Remove each monthly file (and the portal's original beside it) whose month tab is really in
    the workbook just written - read back from disk, never assumed.  Empty folders go too."""
    from openpyxl import load_workbook
    from engines.excel_files import ORIGINALS
    try:
        wb = load_workbook(target, read_only=True)
        try:
            inside = {n.upper() for n in wb.sheetnames}
        finally:
            wb.close()
    except Exception:
        return 0
    removed = 0
    for period, path in folded:
        if period.upper() not in inside:
            continue
        for p in (path, path.parent / ORIGINALS / (path.stem + ".xls")):
            try:
                p.unlink()
                removed += p == path
            except FileNotFoundError:
                pass
            except OSError:
                pass                           # e.g. open in Excel: it is folded in again next run
    for folder in (folded[0][1].parent / ORIGINALS, folded[0][1].parent, folded[0][1].parent.parent) if folded else ():
        try:
            folder.rmdir()                     # only when empty
        except OSError:
            pass
    return removed
