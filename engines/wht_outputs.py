"""
engines/wht_outputs.py - file writers for the WHT (Withholding Tax)
Receivable module.

RECONSTRUCTED from ura_outputs.pyc (Python 3.13 bytecode, APP_VERSION
'3.0'); the original source was not available.  What is exact and what
is re-implemented:

  EXACT (translated instruction by instruction from the bytecode)
    MonthResult, month_sequence, ID_COL / AMOUNT_COL / CHECKBOX_COL,
    flatten_columns, is_amount_column, amount_columns, id_columns,
    _id_text, clean_frame, align_columns, check_period, NAVY / GOLD colours.
    (as_number was widened in 2026-10 to read (500), UGX and /= amounts.  The
    unused password and config.ini code was removed - the suite stores
    passwords only through core/store.py.)
  RE-IMPLEMENTED to the layout the original describes (Summary + Combined
  + one sheet per month; navy header, gold accents, zebra rows, totals row,
  frozen header, auto-filter, #,##0.00 amounts; landscape PDF table with
  'Page x of y'; merged page-copy PDF with a bookmark per month):
    _write_sheet, _run_meta, summary_frame, save_excel, save_month_excels,
    save_csv, _column_widths, save_table_pdf, merge_pdfs.
  Column names, sheet names and file names follow the strings found in
  the bytecode.  Test against one real month before trusting the output.
"""
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

APP_NAME = "URA WHT Extractor"
APP_VERSION = "3.0"
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]


# ====================================================================== data
@dataclass
class MonthResult:
    year: int
    month: int
    df: object = None                 # pandas DataFrame or None
    status: str = "pending"           # pending | ok | empty | error | skipped
    message: str = ""
    pdf_bytes: Optional[bytes] = None # page copy as PDF
    html: Optional[str] = None        # page copy as HTML
    screenshot: Optional[bytes] = None
    seconds: float = 0.0
    warnings: List[str] = field(default_factory=list)

    @property
    def label(self) -> str:           # '2026-06'
        return f"{self.year:04d}-{self.month:02d}"

    @property
    def title(self) -> str:           # 'June 2026'
        return f"{MONTHS[self.month - 1]} {self.year}"

    @property
    def rows(self) -> int:
        return 0 if self.df is None else len(self.df)

    @property
    def portal_period(self) -> str:   # 'MM/YYYY' as the portal prints it
        return f"{self.month:02d}/{self.year:04d}"


def month_sequence(y1, m1, y2, m2):
    """Inclusive list of (year, month) from the first pair to the second."""
    out, y, m = [], y1, m1
    while (y, m) <= (y2, m2):
        out.append((y, m))
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


# ============================================================ table clean-up
ID_COL = re.compile(r"\b(sl|s/?n|sr\.?\s*no\.?|no|nos|num|number|id|tin|code|cert\w*|"
                    r"prn|ref\w*|period)\b", re.I)
AMOUNT_COL = re.compile(r"amount|gross|withheld|tax|value|total", re.I)
CHECKBOX_COL = re.compile(r"^\s*(select|select\s*all|check)?\s*$", re.I)


def flatten_columns(df):
    """pandas.read_html gives a MultiIndex when a header spans two rows
    (the URA grid does: 'Return/Assessment' over 'Number'). Collapse it."""
    import pandas as pd
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join(dict.fromkeys(
            str(x).strip() for x in tup if str(x).strip()
            and not str(x).startswith("Unnamed"))) for tup in df.columns]
    else:
        df.columns = [str(c).strip() for c in df.columns]
    df.columns = [re.sub(r"\s+", " ", c) if c and not c.startswith("Unnamed") else ""
                  for c in df.columns]
    return df


def is_amount_column(name, series=None):
    """Money if the header says so and the values are mostly numbers."""
    import pandas as pd
    name = str(name)
    if ID_COL.search(name) and not AMOUNT_COL.search(name):
        return False
    if not AMOUNT_COL.search(name):
        return False
    if series is None or len(series) == 0:
        return True
    num = pd.to_numeric(series.astype(str).str.replace(",", "", regex=False).str.strip(),
                        errors="coerce")
    return num.notna().sum() >= max(1, int(len(series) * 0.6))


def amount_columns(df):
    return [c for c in df.columns if is_amount_column(c, df[c])]


def id_columns(df):
    return [c for c in df.columns
            if ID_COL.search(str(c)) and not is_amount_column(c, df[c])]


def as_number(value):
    """Amount as a number, or None.  Understands 1,000,000.50, -500, (500) for a
    negative amount, and the currency marks 'UGX' / 'Shs' / '/=' (finding L6)."""
    try:
        s = str(value).replace(",", "").strip()
        s = re.sub(r"(?i)^(ugx|ushs?|shs?)\.?\s*", "", s)
        s = re.sub(r"\s*/=\s*$", "", s)
        negative = s.startswith("(") and s.endswith(")")
        if negative:
            s = s[1:-1].strip()
        if s in ("", "nan", "None", "-"):
            return None
        n = float(s)
        return -n if negative else n
    except (TypeError, ValueError):
        return None


def _id_text(value):
    """'1000025739.0' -> '1000025739'; NaN -> ''."""
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in ("nan", "none"):
        return ""
    if re.fullmatch(r"-?\d+\.0+", s):
        s = s.split(".")[0]
    return s


def clean_frame(df):
    """Tidy a raw scraped grid:
    - flatten two-row headers
    - drop the checkbox column and header-less empty columns
    - KEEP real columns that happen to be empty (e.g. 'TCC No' when no
      certificate has been generated yet) so every month lines up
    - identifiers become clean text, blanks become ''"""
    import pandas as pd
    df = flatten_columns(df.copy())
    keep = []
    for c in df.columns:
        if CHECKBOX_COL.match(c):
            continue
        if c == "" and df[c].isna().all():
            continue
        keep.append(c)
    df = df.loc[:, keep]
    df.columns = [c if c else f"Column {i}" for i, c in enumerate(df.columns, 1)]
    df = df.dropna(axis=0, how="all").reset_index(drop=True)
    if len(df):
        first = df.iloc[:, 0].astype(str).str.strip().str.lower()
        df = df[~first.isin(["total", "grand total"])].reset_index(drop=True)
    for c in id_columns(df):
        df[c] = df[c].map(_id_text)
    for c in df.columns:
        if c in amount_columns(df):
            df[c] = df[c].map(lambda v: "" if pd.isna(v) else v)
    return df


def align_columns(frames):
    """Union of columns, in first-seen order, applied to every frame."""
    order = []
    for df in frames:
        for c in df.columns:
            if c not in order:
                order.append(c)
    return [df.reindex(columns=order) for df in frames], order


def check_period(df, expected):
    """Return the rows whose Period column is NOT the expected 'MM/YYYY'.
    None if there is no period column to check."""
    col = next((c for c in df.columns if "period" in str(c).lower()), None)
    if col is None:
        return None
    import pandas as pd
    vals = df[col].map(lambda v: "" if pd.isna(v) else str(v).strip())
    return df[~vals.isin([expected, ""])]


# ================================================================== Excel
NAVY = "1F4E79"
GOLD = "F5B400"


def _write_sheet(ws, df, title, subtitle=None, amount_cols=None, totals=True,
                 meta_lines=()):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from engines.sheet_safety import as_text_cell
    from openpyxl.utils import get_column_letter
    amount_cols = list(amount_cols or [])
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color=NAVY)
    row = 2
    if subtitle:
        ws.cell(row=row, column=1, value=subtitle).font = Font(name="Arial", size=10,
                                                               italic=True, color="66737F")
        row += 1
    for line in meta_lines:
        ws.cell(row=row, column=1, value=line).font = Font(name="Arial", size=9, color="66737F")
        row += 1
    row += 1
    header_row = row
    for j, col in enumerate(df.columns, 1):
        c = ws.cell(row=row, column=j, value=str(col))
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[row].height = 30
    for i, (_, rec) in enumerate(df.iterrows()):
        row += 1
        zebra = PatternFill("solid", fgColor="F2F6FA") if i % 2 else None
        for j, col in enumerate(df.columns, 1):
            v = rec[col]
            if col in amount_cols:
                v = as_number(v)
            elif v is None or (isinstance(v, float) and v != v):
                v = ""
            c = ws.cell(row=row, column=j, value=v)
            as_text_cell(c)                       # portal text never becomes a formula (L2)
            c.font = Font(name="Arial", size=10)
            c.border = border
            if zebra:
                c.fill = zebra
            if col in amount_cols:
                c.number_format = "#,##0.00"
                c.alignment = Alignment(horizontal="right")
    last_data = row
    if totals and amount_cols and len(df):
        row += 1
        ws.cell(row=row, column=1, value="TOTAL").font = Font(name="Arial", bold=True)
        for j, col in enumerate(df.columns, 1):
            c = ws.cell(row=row, column=j)
            c.fill = PatternFill("solid", fgColor="FFF3CC")
            c.border = border
            if col in amount_cols:
                L = get_column_letter(j)
                c.value = f"=SUM({L}{header_row + 1}:{L}{last_data})"
                c.number_format = "#,##0.00"
                c.font = Font(name="Arial", bold=True)
    for j, col in enumerate(df.columns, 1):
        longest = max([len(str(col))] + [len(str(v)) for v in df[col].tolist()[:500]])
        ws.column_dimensions[get_column_letter(j)].width = min(45, max(10, longest + 2))
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    if len(df):
        ws.auto_filter.ref = (f"A{header_row}:"
                              f"{get_column_letter(len(df.columns))}{last_data}")


def _run_meta(tin="", gen_label="", months=None):
    lines = [f"Extracted: {datetime.now():%d %b %Y %H:%M}  -  {APP_NAME} {APP_VERSION}"]
    if tin:
        masked = tin[:4] + "*" * max(0, len(tin) - 6) + tin[-2:] if len(tin) > 6 else tin
        lines.append(f"Account: {masked}")
    if gen_label:
        lines.append(f"Generated for: {gen_label}")
    if months:
        lines.append(f"Requested: {len(months)} month(s)")
    return lines


def summary_frame(results):
    """One line per requested month, including empty and failed ones."""
    import pandas as pd
    rows = []
    money_cols = []
    for r in results:
        if r.df is not None and len(r.df):
            for c in amount_columns(r.df):
                if c not in money_cols:
                    money_cols.append(c)
    for r in results:
        status = {"ok": "OK", "empty": "No records", "error": "FAILED",
                  "skipped": "Skipped", "pending": "Not run"}.get(r.status, r.status)
        line = {"Month": r.title, "Period": r.portal_period, "Rows": r.rows,
                "Status": status, "Note": r.message or "; ".join(r.warnings)}
        for c in money_cols:
            tot = 0.0
            if r.df is not None and c in r.df.columns:
                tot = sum(x for x in (as_number(v) for v in r.df[c]) if x is not None)
            line[f"Total {c}"] = round(tot, 2)
        rows.append(line)
    return pd.DataFrame(rows)


def save_excel(results, path, tin="", gen_label=""):
    """Summary + Combined + one sheet per month with data."""
    import pandas as pd
    from openpyxl import Workbook
    results = list(results)
    with_data = [r for r in results if r.df is not None and len(r.df)]
    frames, order = align_columns([r.df for r in with_data]) if with_data else ([], [])
    meta = _run_meta(tin, gen_label, results)

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    summ = summary_frame(results)
    _write_sheet(ws, summ, "Withholding Tax Credit - Summary by month",
                 subtitle=f"{len(with_data)} month(s) with data   |   "
                          f"{len(results)} month(s) requested",
                 amount_cols=[c for c in summ.columns if c.startswith("Total ")],
                 meta_lines=meta)
    if with_data:
        tagged = []
        for r, df in zip(with_data, frames):
            d = df.copy()
            d.insert(0, "Month", r.title)
            tagged.append(d)
        all_df = pd.concat(tagged, ignore_index=True)
        ws = wb.create_sheet("Combined")
        _write_sheet(ws, all_df, "Withholding Tax Credit - all months",
                     subtitle=f"Period: {with_data[0].title} to {with_data[-1].title}",
                     amount_cols=amount_columns(all_df), meta_lines=meta)
        for r, df in zip(with_data, frames):
            ws = wb.create_sheet(r.label)
            _write_sheet(ws, df, f"Withholding Tax Credit - {r.title}",
                         subtitle=f"Period: {r.portal_period}",
                         amount_cols=amount_columns(df))
    os.makedirs(os.path.dirname(str(path)) or ".", exist_ok=True)
    wb.save(path)
    return path


def save_month_excels(results, folder, prefix="WHT"):
    """One small workbook per month with data. Returns the paths."""
    from openpyxl import Workbook
    paths = []
    os.makedirs(folder, exist_ok=True)
    for r in results:
        if r.df is None or not len(r.df):
            continue
        wb = Workbook()
        ws = wb.active
        ws.title = r.label
        _write_sheet(ws, r.df, f"Withholding Tax Credit - {r.title}",
                     subtitle=f"Period: {r.portal_period}", amount_cols=amount_columns(r.df))
        p = os.path.join(folder, f"{prefix}_{r.label}.xlsx")
        wb.save(p)
        paths.append(p)
    return paths


def save_csv(results, path):
    """One flat CSV of every row, with a Month column first."""
    import pandas as pd
    with_data = [r for r in results if r.df is not None and len(r.df)]
    if not with_data:
        return None
    frames, _ = align_columns([r.df for r in with_data])
    tagged = []
    for r, df in zip(with_data, frames):
        d = df.copy()
        d.insert(0, "Month", r.title)
        tagged.append(d)
    from engines.sheet_safety import csv_text
    flat = pd.concat(tagged, ignore_index=True)
    flat = flat.apply(lambda col: col.map(csv_text) if col.dtype == object else col)
    flat.to_csv(path, index=False, encoding="utf-8-sig")
    return path


# ==================================================================== PDF
def _fmt_amount(v):
    n = as_number(v)
    return "" if n is None else f"{n:,.2f}"


def _column_widths(df, avail, char_pt=4.3, min_pt=12.0):
    """Points per column: never narrower than the longest single word (so
    numbers and TINs do not wrap mid-value), then share the rest by how much
    text each column carries."""
    texts = {c: [str(c)] + [str(v) for v in df[c].tolist()] for c in df.columns}
    longest_word = {c: max(len(w) for t in texts[c] for w in (t.split() or [""]))
                    for c in df.columns}
    longest_text = {c: max(len(t) for t in texts[c]) for c in df.columns}
    mins = {c: max(min_pt, longest_word[c] * char_pt + 6) for c in df.columns}
    total_min = sum(mins.values())
    if total_min >= avail:
        scale = avail / total_min
        return [mins[c] * scale for c in df.columns]
    spare = avail - total_min
    weight = {c: max(0, longest_text[c] - longest_word[c]) for c in df.columns}
    tot_w = sum(weight.values()) or 1
    return [mins[c] + spare * weight[c] / tot_w for c in df.columns]


def save_table_pdf(results, path, tin="", gen_label=""):
    """Typeset table, landscape A4, formatted amounts, totals, page x of y."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.platypus import (PageBreak, Paragraph, SimpleDocTemplate,
                                    Spacer, Table, TableStyle)

    class NumberedCanvas(rl_canvas.Canvas):
        """'Page x of y' footer."""
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self._saved = []

        def showPage(self):
            self._saved.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            n = len(self._saved)
            for state in self._saved:
                self.__dict__.update(state)
                self.setFont("Helvetica", 8)
                self.setFillColor(colors.HexColor("#777777"))
                w, h = self._pagesize
                self.drawRightString(w - 10 * mm, 6 * mm, f"Page {self._pageNumber} of {n}")
                self.drawString(10 * mm, 6 * mm,
                                f"{APP_NAME}  -  extracted {datetime.now():%d %b %Y %H:%M}")
                super().showPage()
            super().save()

    pagesize = landscape(A4)
    doc = SimpleDocTemplate(str(path), pagesize=pagesize, leftMargin=10 * mm,
                            rightMargin=10 * mm, topMargin=10 * mm, bottomMargin=12 * mm,
                            title="URA Withholding Tax Credit", author=APP_NAME)
    ss = getSampleStyleSheet()
    cell = ParagraphStyle("cell", parent=ss["Normal"], fontName="Helvetica",
                          fontSize=7.5, leading=9)
    head = ParagraphStyle("head", parent=cell, fontName="Helvetica-Bold",
                          textColor=colors.white)
    small = ParagraphStyle("small", parent=ss["Normal"], fontSize=8,
                           textColor=colors.HexColor("#555555"))
    avail = pagesize[0] - 20 * mm
    story = [Paragraph("URA Withholding Tax Credit", ss["Title"])]
    for line in _run_meta(tin, gen_label, list(results)):
        story.append(Paragraph(line, small))
    story.append(Spacer(1, 6))

    def esc(s):
        return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

    first = True
    for r in results:
        if r.df is None or not len(r.df):
            continue
        if not first:
            story.append(PageBreak())
        first = False
        df = r.df
        money = amount_columns(df)
        story.append(Paragraph(f"{r.title}  (period {r.portal_period}, {len(df)} rows)",
                               ss["Heading3"]))
        data = [[Paragraph(esc(c), head) for c in df.columns]]
        for _, rec in df.iterrows():
            data.append([Paragraph(esc(_fmt_amount(rec[c]) if c in money else rec[c]), cell)
                         for c in df.columns])
        if money:
            tot = []
            for c in df.columns:
                if c in money:
                    s = sum(x for x in (as_number(v) for v in df[c]) if x is not None)
                    tot.append(Paragraph(f"<b>{s:,.2f}</b>", cell))
                else:
                    tot.append(Paragraph("<b>TOTAL</b>" if c == df.columns[0] else "", cell))
            data.append(tot)
        t = Table(data, colWidths=_column_widths(df, avail), repeatRows=1)
        style = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + NAVY)),
                 ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
                 ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                 ("ROWBACKGROUNDS", (0, 1), (-1, -1),
                  [colors.white, colors.HexColor("#F2F6FA")])]
        for j, c in enumerate(df.columns):
            if c in money:
                style.append(("ALIGN", (j, 1), (j, -1), "RIGHT"))
        if money:
            style.append(("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#FFF3CC")))
        t.setStyle(TableStyle(style))
        story.append(t)
    if first:
        story.append(Paragraph("No records were found for the requested months.", ss["Normal"]))
    doc.build(story, canvasmaker=NumberedCanvas)
    return path


def merge_pdfs(results, path):
    """Join the page-copy PDFs, with a bookmark per month."""
    from pypdf import PdfReader, PdfWriter
    import io
    w = PdfWriter()
    any_page = False
    for r in results:
        if not r.pdf_bytes:
            continue
        reader = PdfReader(io.BytesIO(r.pdf_bytes))
        start = len(w.pages)
        for p in reader.pages:
            w.add_page(p)
        w.add_outline_item(r.title, start)
        any_page = True
    if not any_page:
        return None
    with open(path, "wb") as fh:
        w.write(fh)
    return path


def safe_prefix(text):
    text = re.sub(r"[^A-Za-z0-9_\-]+", "_", str(text or "")).strip("_")
    return text or "WHT"
