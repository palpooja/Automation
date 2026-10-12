"""10.13 (user, 8 Oct 2026): WHT certificates (Payee) checked against the Income Tax ledger's Level 3 lines.
10.15: the live layout (capture of 8 Oct 2026) - the line's reference is the agent's RETURN number.
Invented data only."""
import pandas as pd
from openpyxl import load_workbook

from engines import excel_files, wht_crosscheck as wc
from engines.wht_outputs import MonthResult, save_excel


def month(y, m, rows):
    r = MonthResult(y, m)
    if rows:
        r.df = pd.DataFrame(rows, columns=["Sr.No.", "Agent TIN", "Agent Name", "Period (Month/Year)",
                                           "Return/Assessment Number", "Nature Of Payment", "Gross Amount",
                                           "Amount Withheld", "Creditable/Non Creditable", "TCC No."])
        r.status = "ok"
    else:
        r.status = "empty"
    return r


def ledger(folder, name, lines):
    body = "".join("<tr>" + "".join(f"<td>{v}</td>" for v in ln) + "</tr>" for ln in lines)
    html = ('<table><tr><td colspan="6">Tax Position as at 08/10/2026 10:00:00</td></tr>'
            "<tr><th>Sr No.</th><th>Payment Basis</th><th>Effective Date</th><th>Reference</th>"
            "<th>Amount</th><th>Adjusted Against Period</th></tr>" + body + "</table>")
    excel_files.html_to_xlsx(html, folder / name)


def test_certificates_found_by_tcc_or_amount_and_gaps_listed(tmp_path):
    led = tmp_path / "Ledgers" / "1000000001"
    led.mkdir(parents=True)
    ledger(led, "Income Tax_L3_TaxCredit_JUL-2025 to JUN-2026_1500000.xlsx", [
        ["1", "WHT Credit", "15-AUG-2025", "TCC000000000101", "600,000", "JUL-2025 to JUN-2026"],
        ["2", "WHT Credit", "15-NOV-2025", "TCC000000000999", "300,000", "JUL-2025 to JUN-2026"],
        ["3", "WHT Credit", "15-JAN-2026", "", "450,000", "JUL-2025 to JUN-2026"]])
    months = [
        month(2025, 7, [[1, "1000000201", "AGENT ONE", "07/2025", "WT0000001", "Services", "10,000,000",
                         "600,000", "Creditable", "TCC000000000101"]]),
        month(2025, 8, []),
        month(2025, 12, [[1, "1000000202", "AGENT TWO", "12/2025", "WT0000002", "Services", "7,500,000",
                          "450,000", "Creditable", ""],
                         [2, "1000000203", "AGENT THREE", "12/2025", "WT0000003", "Services", "2,000,000",
                          "120,000", "Creditable", "TCC000000000777"]])]
    certs, left, files, _ = wc.crosscheck(months, led)
    assert list(certs["Found in Income Tax ledger"]) == ["Yes - TCC No.", "Yes - amount", "No"]
    assert [ln["text"].split(" | ")[3] for ln in left] == ["TCC000000000999"], "ledger credit with no certificate"
    book = tmp_path / "WHT.xlsx"
    save_excel(months, str(book), tin="1000000001", gen_label="Payee (Deductee)")
    wc.add_sheet(book, (certs, left, files))
    wb = load_workbook(book)
    assert wb.sheetnames[:2] == ["Summary", "Check vs ledger"]
    first = wb["Check vs ledger"]["A1"].value
    assert first.startswith("3 certificate(s); 2 found in the Income Tax ledger, 1 not found; 1 ledger")


def test_without_income_tax_level3_the_sheet_says_how_to_get_it(tmp_path):
    assert wc.crosscheck([month(2025, 7, [])], tmp_path / "Ledgers" / "1000000001") is None
    book = tmp_path / "WHT.xlsx"
    save_excel([month(2025, 7, [])], str(book), tin="1000000001", gen_label="Payee (Deductee)")
    wc.add_sheet(book, None)
    assert "Not checked" in load_workbook(book)["Check vs ledger"]["A1"].value


def test_certificates_step_puts_missing_ones_on_the_review_page(tmp_path):
    from pathlib import Path
    from modules.wht_receivable import WhtReceivableModule

    class Ctx:
        output_root = tmp_path
        logs, items = [], []
        def log(self, m, kind=""): self.logs.append(m)
        def record(self, item, status, detail="", file=""): self.items.append((item, status, detail))
    led = tmp_path / "Ledgers" / "1000000001"
    led.mkdir(parents=True)
    ledger(led, "Income Tax_L3_TaxCredit_JUL-2025 to JUN-2026_600000.xlsx",
           [["1", "WHT Credit", "15-AUG-2025", "TCC000000000101", "600,000", "JUL-2025 to JUN-2026"]])
    months = [month(2025, 12, [[1, "1000000203", "AGENT THREE", "12/2025", "WT0000003", "Services",
                                "2,000,000", "120,000", "Creditable", "TCC000000000777"]])]
    book = tmp_path / "WHT.xlsx"
    save_excel(months, str(book), tin="1000000001", gen_label="Payee (Deductee)")
    ctx = Ctx()
    WhtReceivableModule._check_vs_ledger(ctx, "1000000001", months, str(book))
    assert ("WHT Payee December 2025 - TCC TCC000000000777", "check") in [(i, s) for i, s, _ in ctx.items]
    assert any(s == "check" and i.startswith("Income Tax ledger WHT credit") for i, s, _ in ctx.items)
    assert any("0 of 1 certificate(s) found" in m for m in ctx.logs), ctx.logs


def test_live_layout_matched_by_return_number_from_the_income_tax_workbook(tmp_path):
    """Income Tax Level 3-2 as captured on 8 Oct 2026: 'WHT Credit' lines carry the withholding agent's return
    number (MTO...) as Payment Registration Number; 'ASYCUDA Credit' is a customs credit.  Since 10.15 the
    lines are in Ledgers\\<TIN>\\Income Tax.xlsx (the yearly files are folded in and removed)."""
    from engines import ledger_taxhead as th
    from tests.auto.test_ledger_tax_head_file import level3, line
    led = tmp_path / "Ledgers" / "1000000001"
    year = "JUL-2025 to JUN-2026"
    level3(th.months_folder(led, "Income Tax"), "Income Tax", year, [
        line(1, "Payment (PRN)", "2250000000401", "2,500,000", "NA"),
        line(2, "WHT Credit", "MTO000000000501", "450,000", "NA"),
        line(3, "WHT Credit", "MTO000000000502", "750,000", "NA"),
        line(4, "ASYCUDA Credit", "C 0000 2026", "400,000", "NA")])
    th.update_tax_head(led, "Income Tax")
    assert not list(th.months_folder(led, "Income Tax").glob("*.xlsx")), "only the workbook is left"
    months = [month(2025, 11, [[1, "1000000201", "AGENT ONE", "11/2025", "MTO000000000501", "Services",
                                "7,500,000", "450,000", "Creditable", "TCC000000000101"],
                               [2, "1000000202", "AGENT TWO", "11/2025", "MTO000000000999", "Services",
                                "1,000,000", "60,000", "Creditable", "TCC000000000102"]])]
    certs, left, files, others = wc.crosscheck(months, led)
    assert files == ["Income Tax.xlsx"]
    assert list(certs["Found in Income Tax ledger"]) == ["Yes - return number", "No"]
    assert [ln["ref"] for ln in left] == ["MTO000000000502"], "ledger WHT credit with no certificate"
    assert [ln["basis"] for ln in others] == ["ASYCUDA Credit"], "customs credit listed apart, never a 'check'"
    book = tmp_path / "WHT.xlsx"
    save_excel(months, str(book), tin="1000000001", gen_label="Payee (Deductee)")
    wc.add_sheet(book, (certs, left, files, others))
    cells = [c for r in load_workbook(book)["Check vs ledger"].iter_rows(values_only=True) for c in r if c]
    assert any(str(c).startswith("Other credits in the Income Tax ledger") for c in cells)
