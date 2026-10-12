"""engines/excel_files.py - make every ledger 'Excel' file a real Excel workbook.

Live run, 3 Oct 2026: Excel refused every ledger file saved from the portal's
'Export To Excel' button ("file format or file extension is not valid").  The
portal's exportTableToExcel() sends an HTML table and names it Level-2.xls; the
program saved it as <tax>_L2.xlsx.  Excel checks the contents against the name
and refuses.

What this module does:
* `kind(path)`       - what a file really is, from its first bytes (not its name);
* `html_to_xlsx()`   - the portal's table -> a real .xlsx with exactly the same text:
                       header cells bold, merged where the portal merges them;
                       only amounts become numbers (identifiers, TINs, PRNs, return
                       numbers, dates and codes with leading zeros stay text);
                       every cell is read back and compared before the file is kept;
* `save_portal_excel()` - used for each download; the portal's own file is kept in
                       portal_original/ beside it (named .xls, as the portal names it);
* `repair_folder()`  - converts files saved by earlier versions in place, keeping
                       their date (so 'saved on an earlier day' still works).
"""
import os
import re
from pathlib import Path

ZIP_MAGIC = b"PK\x03\x04"                       # .xlsx is a zip file
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"   # old binary .xls
ORIGINALS = "portal_original"

_AMOUNT = re.compile(r"-?\d{1,3}(,\d{3})+(\.\d+)?|-?\d+\.\d+")
_BRACKETED = re.compile(r"\((\d{1,3}(,\d{3})*(\.\d+)?|\d+(\.\d+)?)\)")
_SHORT_INT = re.compile(r"-?(0|[1-9]\d{0,8})")   # up to 9 digits, no leading zero


def kind(path) -> str:
    """'xlsx' | 'xls' | 'html' | 'other' - judged by the bytes inside the file."""
    try:
        with open(path, "rb") as fh:
            head = fh.read(4096)
    except OSError:
        return "other"
    if head.startswith(ZIP_MAGIC):
        return "xlsx"
    if head.startswith(OLE_MAGIC):
        return "xls"
    text = head.decode("utf-8", errors="ignore").lower()
    if "<table" in text or "<html" in text or "<tr" in text:
        return "html"
    return "other"


def cell_value(text):
    """The value to store for one cell's text: a number only for an amount."""
    t = " ".join(str(text).split())
    if t == "":
        return None
    if _AMOUNT.fullmatch(t):
        n = float(t.replace(",", ""))
        return int(n) if n.is_integer() and "." not in t else n
    m = _BRACKETED.fullmatch(t)
    if m:
        n = -float(m.group(1).replace(",", ""))
        return int(n) if n.is_integer() and "." not in t else n
    if _SHORT_INT.fullmatch(t):
        return int(t)
    return t                                    # identifiers, dates, codes, '007', 10+ digit numbers


def _grid(table):
    """Rows of (text, is_header) with colspan/rowspan laid out as on screen, plus merges."""
    rows = table.xpath("./tr|./thead/tr|./tbody/tr|./tfoot/tr")
    grid, merges, carry = [], [], {}           # carry[(r, c)] = filled by a rowspan above
    for r, tr in enumerate(rows):
        out, c = [], 0
        cells = tr.xpath("./th|./td")
        for cell in cells:
            while (r, c) in carry:
                out.append(carry.pop((r, c)))
                c += 1
            text = " ".join(cell.text_content().split())
            head = cell.tag.lower() == "th"
            cs = max(1, int(cell.get("colspan", "1") or 1))
            rs = max(1, int(cell.get("rowspan", "1") or 1))
            if cs > 1 or rs > 1:
                merges.append((r, c, r + rs - 1, c + cs - 1))
            for dc in range(cs):
                out.append((text if dc == 0 else "", head))
                for dr in range(1, rs):
                    carry[(r + dr, c + dc)] = ("", head)
            c += cs
        while (r, c) in carry:
            out.append(carry.pop((r, c)))
            c += 1
        grid.append(out)
    return grid, merges


def html_to_xlsx(html_text: str, dest) -> int:
    """Write the portal's table(s) to a real .xlsx at dest.  Returns the number of rows.
    Raises ValueError when there is no table (e.g. a 'session expired' page), so a
    broken download is never kept as a finished file."""
    from lxml import html as LH
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    doc = LH.fromstring(html_text)
    tables = doc.xpath("//table[not(ancestor::table)]")
    if not tables:
        raise ValueError("the portal's file contains no table")
    dest = Path(dest)
    wb = Workbook()
    wb.remove(wb.active)
    expected = []
    total = 0
    for i, table in enumerate(tables, 1):
        grid, merges = _grid(table)
        ws = wb.create_sheet("Sheet1" if i == 1 else f"Sheet{i}")
        width = {}
        for r, row in enumerate(grid, 1):
            for c, (text, head) in enumerate(row, 1):
                v = cell_value(text)
                cell = ws.cell(row=r, column=c, value=v)
                if isinstance(v, str):
                    cell.data_type = "s"            # never a formula, whatever the text
                if head:
                    cell.font = Font(bold=True)
                if isinstance(v, (int, float)) and "," in text:
                    cell.number_format = "#,##0.00" if isinstance(v, float) else "#,##0"
                width[c] = max(width.get(c, 0), len(text))
                expected.append((ws.title, r, c, v))
        for r1, c1, r2, c2 in merges:
            ws.merge_cells(start_row=r1 + 1, start_column=c1 + 1, end_row=r2 + 1, end_column=c2 + 1)
        for c, w in width.items():
            ws.column_dimensions[get_column_letter(c)].width = min(50, max(8, w + 2))
        total += len(grid)
    tmp = dest.with_name("~" + dest.stem + ".tmp.xlsx")       # openpyxl needs an .xlsx name
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb.save(tmp)
    check = load_workbook(tmp)                   # read every cell back before keeping the file
    for sheet, r, c, v in expected:
        got = check[sheet].cell(row=r, column=c).value
        if got != v and not (v is None and got in (None, "")):
            tmp.unlink(missing_ok=True)
            raise ValueError(f"conversion check failed at {sheet}!R{r}C{c}")
    os.replace(tmp, dest)
    return total


def compare_cells(a, b, limit=5, ignore_text=()):
    """(same, differences) for two workbooks: sheet names and every cell value (formatting is
    ignored).  Empty and None count as equal.  At most `limit` differences are listed.  A cell pair
    where either value contains one of `ignore_text` is skipped (e.g. 'Tax Position as at', the time
    of the download, which differs every time)."""
    from openpyxl import load_workbook
    wa, wb = load_workbook(a, read_only=True), load_workbook(b, read_only=True)
    diffs = []
    try:
        if wa.sheetnames != wb.sheetnames:
            diffs.append(f"sheets {wa.sheetnames} vs {wb.sheetnames}")
        for name in wa.sheetnames:
            if name not in wb.sheetnames:
                continue
            ra = list(wa[name].iter_rows(values_only=True))
            rb = list(wb[name].iter_rows(values_only=True))
            blank = lambda v: v is None or v == ""
            while ra and all(blank(v) for v in ra[-1]):
                ra.pop()
            while rb and all(blank(v) for v in rb[-1]):
                rb.pop()
            if len(ra) != len(rb):
                diffs.append(f"{name}: {len(ra)} rows vs {len(rb)} rows")
            for i, (x, y) in enumerate(zip(ra, rb), 1):
                for j in range(max(len(x), len(y))):
                    vx = x[j] if j < len(x) else None
                    vy = y[j] if j < len(y) else None
                    if blank(vx) and blank(vy):
                        continue
                    if ignore_text and any(isinstance(v, str) and t in v for v in (vx, vy) for t in ignore_text):
                        continue
                    if vx != vy:
                        diffs.append(f"{name}!R{i}C{j + 1}: {vx!r} vs {vy!r}")
                        if len(diffs) >= limit:
                            return False, diffs
    finally:
        wa.close()
        wb.close()
    return not diffs, diffs


def _keep_original(raw: bytes, dest: Path) -> Path:
    keep = dest.parent / ORIGINALS / (dest.stem + ".xls")
    keep.parent.mkdir(parents=True, exist_ok=True)
    keep.write_bytes(raw)
    return keep


def _decode(raw: bytes) -> str:
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def save_portal_excel(src, dest) -> str:
    """Turn a downloaded portal file (at src) into dest (.xlsx).  Returns what it was."""
    src, dest = Path(src), Path(dest)
    what = kind(src)
    raw = src.read_bytes()
    if what == "xlsx":
        os.replace(src, dest)
        return "xlsx"
    if what == "html":
        html_to_xlsx(_decode(raw), dest)
        _keep_original(raw, dest)
        src.unlink(missing_ok=True)
        return "html converted"
    if what == "xls":
        # a genuine old-format workbook: keep it under its true extension
        os.replace(src, dest.with_suffix(".xls"))
        raise ValueError("the portal sent an old-format .xls workbook - saved as "
                         f"{dest.with_suffix('.xls').name}; it is not converted")
    raise ValueError("the portal's file is not a table or a workbook")


def repair_folder(root) -> int:
    """Convert every '.xlsx' under root that is really an HTML table.  Keeps each file's
    date and a copy of the original in portal_original/.  Returns how many were repaired."""
    root = Path(root)
    if not root.exists():
        return 0
    repaired = 0
    for path in root.rglob("*.xlsx"):
        if ORIGINALS in path.parts or kind(path) != "html":
            continue
        try:
            st = path.stat()
            raw = path.read_bytes()
            html_to_xlsx(_decode(raw), path)
            _keep_original(raw, path)
            os.utime(path, (st.st_atime, st.st_mtime))
            repaired += 1
        except ValueError:
            # not convertible: set it aside so it no longer counts as downloaded
            try:
                os.replace(path, path.with_name(path.name + ".broken"))
            except OSError:
                pass
        except OSError:
            continue
    return repaired


def excel_ok(path) -> bool:
    """A file that exists AND really is an Excel workbook."""
    return Path(path).exists() and kind(path) == "xlsx"
