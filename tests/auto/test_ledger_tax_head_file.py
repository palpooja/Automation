"""User, 6 Oct 2026: 'one file per Tax Head, only one table, keep updating data instead of making a
new file every time; make sure it is not taking extra time.'

Each tax head gets Ledgers\\<TIN>\\<Tax head>.xlsx with ONE Excel table (Period, Tax Position as at, then
the portal's Level 3-2 columns); a later run replaces the months it fetched and keeps the others.
The portal's own monthly files go to Ledgers\\<TIN>\\_months\\<Tax head>\\ (needed to resume a run, for the
route proof, and as the portal's evidence).  Level 3 test files are made like the portal's: converted
from the portal table, first line 'Tax Position as at ...' (live log 6 Oct 2026).
"""
import time

from openpyxl import load_workbook

from engines import excel_files, ledger_engine as eng, ledger_taxhead as th

HEAD = ["Sr No.", "Payment Basis", "Effective Date", "Payment Registration Number", "Amount",
        "Adjusted Against Period", "Liability Basis", "Principal", "Penalty", "Interest", "Refund"]


def level3(folder, tax, month, rows, stamp="06/10/2026 05:31:40"):
    """A Level 3 file exactly as the program saves the portal's export (excel_files.html_to_xlsx)."""
    body = "".join("<tr>" + "".join(f"<td>{v}</td>" for v in r) + "</tr>" for r in rows)
    html = (f'<table><tr><td colspan="11">Tax Position as at {stamp}</td></tr>'
            f'<tr><td colspan="11">Period: {month}</td></tr>'
            "<tr>" + "".join(f"<th>{h}</th>" for h in HEAD) + "</tr>" + body + "</table>")
    folder.mkdir(parents=True, exist_ok=True)
    excel_files.html_to_xlsx(html, folder / f"{tax}_L3_{month}.xlsx")


def row(n, prn, amount, month):
    return [n, "Payment (PRN)", "14-SEP-2026", prn, amount, month, "Original Return", amount, 0, 0, 0]


def table_rows(path):
    wb = load_workbook(path)
    assert wb.sheetnames[:2] == ["Summary", "Level 3-2"], wb.sheetnames
    ws = wb["Level 3-2"]
    assert len(ws.tables) == 1, "exactly one table"
    return [list(r) for r in ws.iter_rows(values_only=True)]


def test_one_file_one_table_with_every_month_oldest_first(tmp_path):
    months = th.months_folder(tmp_path, "PAYE")
    level3(months, "PAYE", "JUN-2026", [row(1, "2270000505613", "1,160,500", "JUN-2026")])
    level3(months, "PAYE", "JUL-2025", [row(1, "2260003116559", "1,160,500", "JUL-2025"),
                                        row(2, "2260003116560", "15,000", "JUL-2025")])
    path = th.update_tax_head(tmp_path, "PAYE")
    assert path == tmp_path / "PAYE.xlsx"
    rows = table_rows(path)
    assert rows[0] == ["Period", "Tax Position as at"] + HEAD
    assert [r[0] for r in rows[1:]] == ["JUL-2025", "JUL-2025", "JUN-2026"]
    assert rows[1][1] == "06/10/2026 05:31:40"
    assert rows[1][5] == "2260003116559" and rows[1][6] == 1160500, "PRN kept as text, amount as number"


def test_next_run_updates_the_same_file(tmp_path):
    months = th.months_folder(tmp_path, "PAYE")
    level3(months, "PAYE", "JUN-2026", [row(1, "2270000505613", "1,160,500", "JUN-2026")])
    level3(months, "PAYE", "MAY-2026", [row(1, "2260026009392", "1,160,500", "MAY-2026")])
    th.update_tax_head(tmp_path, "PAYE")
    # next run: MAY fetched again with a new payment, JUL-2026 new; JUN's month file was folded in and removed
    assert not (months / "PAYE_L3_JUN-2026.xlsx").exists()
    level3(months, "PAYE", "MAY-2026", [row(1, "2260026009392", "1,160,500", "MAY-2026"),
                                        row(2, "2260026009999", "500", "MAY-2026")], stamp="20/10/2026 09:00:00")
    level3(months, "PAYE", "JUL-2026", [row(1, "2270002000001", "1,200,000", "JUL-2026")])
    before = sorted(p.name for p in tmp_path.glob("*.xlsx"))
    th.update_tax_head(tmp_path, "PAYE")
    assert sorted(p.name for p in tmp_path.glob("*.xlsx")) == before, "no new file - the same one updated"
    rows = table_rows(tmp_path / "PAYE.xlsx")
    assert [(r[0], r[5]) for r in rows[1:]] == [("MAY-2026", "2260026009392"), ("MAY-2026", "2260026009999"),
                                                ("JUN-2026", "2270000505613"), ("JUL-2026", "2270002000001")]
    assert rows[1][1] == "20/10/2026 09:00:00", "a re-fetched month replaces its old rows"
    wb = load_workbook(tmp_path / "PAYE.xlsx")
    assert wb.sheetnames == ["Summary", "Level 3-2", "MAY-2026", "JUN-2026", "JUL-2026"]
    assert "2270000505613" in _cells(wb["JUN-2026"]), "a month tab is kept when its month was not fetched again"
    assert "2260026009999" in _cells(wb["MAY-2026"]), "a re-fetched month replaces its tab"


def test_month_files_from_earlier_versions_are_moved_not_fetched_again(tmp_path):
    level3(tmp_path, "PAYE", "JUN-2026", [row(1, "2270000505613", "1,160,500", "JUN-2026")])  # old place
    moved = th.move_old_month_files(tmp_path, "PAYE")
    target = th.months_folder(tmp_path, "PAYE") / "PAYE_L3_JUN-2026.xlsx"
    assert moved == 1 and target.exists() and not (tmp_path / "PAYE_L3_JUN-2026.xlsx").exists()
    assert eng.month_file_done(tmp_path, "PAYE", "JUN-2026", target)


def test_updating_takes_no_noticeable_time(tmp_path):
    months = th.months_folder(tmp_path, "Value Added Tax")
    for i, m in enumerate(["JUL-2025", "AUG-2025", "SEP-2025", "OCT-2025", "NOV-2025", "DEC-2025",
                           "JAN-2026", "FEB-2026", "MAR-2026", "APR-2026", "MAY-2026", "JUN-2026"]):
        level3(months, "Value Added Tax", m, [row(n, f"22600{i:02d}{n:06d}", "1,000", m) for n in range(1, 16)])
    t0 = time.monotonic()
    th.update_tax_head(tmp_path, "Value Added Tax")
    took = time.monotonic() - t0
    assert len(table_rows(tmp_path / "Value Added Tax.xlsx")) == 1 + 12 * 15
    assert took < 2.0, f"took {took:.2f} s for 12 months x 15 payments"


import pytest  # noqa: E402


@pytest.mark.browser
def test_real_tax_type_routine_fills_the_tax_head_file(page, mock_url, ledger_portal, tmp_path, monkeypatch):
    """process_tax_type on the test copy of the live Ledger form (option C), then the end-of-step update."""
    ledger_portal.reset()
    page.goto(f"{mock_url}{ledger_portal.BASE}loadLevelTwo")
    monkeypatch.setattr(eng, "open_level2", lambda pg, *a, **k: pg.main_frame)        # Level 1 -> 2 not on the test portal
    monkeypatch.setattr(eng, "recover_to_level1", lambda *a, **k: None)
    monkeypatch.setattr(eng, "_back_to_level1_after", lambda *a, **k: None)
    monkeypatch.setattr(eng, "export_level2_to_excel", lambda *a, **k: "portal export")
    monkeypatch.setattr(eng, "SKIP_NIL", False)
    import pandas as pd                                           # Level 2 reading is tested elsewhere
    monkeypatch.setattr(eng, "read_level2_table", lambda fr: pd.DataFrame(
        {"Period": [m for m, _ in ledger_portal.MONTHS], "Payment": [a for _, a in ledger_portal.MONTHS]}))
    eng.ROUTE_STATE.clear()
    old = tmp_path / "PAYE_L3_JUN-2026.xlsx"                      # a month saved by version 10.8
    level3(tmp_path, "PAYE", "JUN-2026", [row(1, "2270000505613", "1,329,127", "JUN-2026")])
    eng.process_tax_type(page, "PAYE", "2026-04", "2026-06", tmp_path, "1000000001", "pw")
    months = th.months_folder(tmp_path, "PAYE")
    assert not old.exists() and (months / "PAYE_L3_JUN-2026.xlsx").exists()
    assert ledger_portal.STATE["level3_requests"][0] == "MAY-2026", "JUN-2026 was already done - not fetched again"
    assert sorted(p.name for p in tmp_path.glob("*.xlsx")) == [], "no monthly file beside the workbooks"
    eng.update_tax_head_files(tmp_path, ["PAYE"])
    rows = table_rows(tmp_path / "PAYE.xlsx")
    assert [r[0] for r in rows[1:]] == ["APR-2026", "MAY-2026", "JUN-2026"]
    wb = load_workbook(tmp_path / "PAYE.xlsx")
    assert wb.sheetnames == ["Summary", "Level 3-2", "APR-2026", "MAY-2026", "JUN-2026"]
    assert not months.exists(), "no Excel file per month left after the step (user, 8 Oct 2026)"
    assert (tmp_path / eng.ROUTE_CHECK_FILE).exists(), "route proof noted in the client folder, not in _months"
    assert sorted(p.name for p in tmp_path.rglob("*.xls*")) == ["PAYE.xlsx"]
    saved = [r for r in eng.RESULTS if r["Section"] == "PAYE ledger" and r["Item"] in ("APR-2026", "MAY-2026")]
    assert saved and all(r["File"] == str(tmp_path / "PAYE.xlsx") for r in saved), \
        "the run summary links each month to the workbook, not to a removed file"
    # a second run fetches nothing: every month is in the workbook
    ledger_portal.reset()
    page.goto(f"{mock_url}{ledger_portal.BASE}loadLevelTwo")
    eng.RESULTS.clear()
    eng.process_tax_type(page, "PAYE", "2026-04", "2026-06", tmp_path, "1000000001", "pw")
    assert ledger_portal.STATE["level3_requests"] == []


# ---------- 8 Oct 2026 (user): 'I don't want an Excel for every month, I want a combined one like WHT
# Receivables.'  The tax-head workbook now holds Summary, Level 3-2 (all months, one table) and one tab
# per month - like the WHT workbook (Summary, Combined, a tab per month); the monthly files are removed
# once they are safely inside it.


def _cells(ws):
    return [v for r in ws.iter_rows(values_only=True) for v in r if v is not None]


def _two_months(tmp_path, tax="PAYE"):
    months = th.months_folder(tmp_path, tax)
    level3(months, tax, "JUN-2026", [row(1, "2270000505613", "1,160,500", "JUN-2026")])
    level3(months, tax, "JUL-2025", [row(1, "2260003116559", "1,160,500", "JUL-2025"),
                                     row(2, "2260003116560", "15,000", "JUL-2025")])
    excel_files._keep_original(b"<table></table>", months / f"{tax}_L3_JUN-2026.xlsx")   # portal_original copy
    return months


def test_workbook_like_wht_summary_combined_and_a_tab_per_month(tmp_path):
    months = _two_months(tmp_path)
    original = load_workbook(months / "PAYE_L3_JUL-2025.xlsx").worksheets[0]
    want = [list(r) for r in original.iter_rows(values_only=True)]
    merged = sorted(str(m) for m in original.merged_cells.ranges)
    th.update_tax_head(tmp_path, "PAYE")
    wb = load_workbook(tmp_path / "PAYE.xlsx")
    assert wb.sheetnames == ["Summary", "Level 3-2", "JUL-2025", "JUN-2026"], "oldest month first"
    tab = wb["JUL-2025"]
    assert [list(r) for r in tab.iter_rows(values_only=True)] == want, "the portal's month exactly as exported"
    assert sorted(str(m) for m in tab.merged_cells.ranges) == merged
    assert tab["D4"].value == "2260003116559" and tab["D4"].data_type == "s", "PRN stays text"
    summary = [list(r) for r in wb["Summary"].iter_rows(values_only=True)]
    assert summary[0][:5] == ["Period", "Status", "Payment (Level 2)", "Level 3 rows", "Level 3 total"]
    assert [(r[0], r[3], r[4]) for r in summary[1:]] == [("JUL-2025", 2, 1175500), ("JUN-2026", 1, 1160500)]
    assert len(wb["Summary"].tables) == 1


def test_month_files_are_removed_once_inside_the_workbook(tmp_path):
    months = _two_months(tmp_path)
    th.update_tax_head(tmp_path, "PAYE")
    assert not months.exists(), "no monthly .xlsx / portal_original .xls left"
    assert not (tmp_path / th.MONTHS_DIR).exists(), "the empty _months folder is tidied away too"
    assert [p.name for p in tmp_path.iterdir()] == ["PAYE.xlsx"]


def test_option_keeps_a_file_per_month(tmp_path):
    months = _two_months(tmp_path)
    th.update_tax_head(tmp_path, "PAYE", keep_month_files=True)
    assert (months / "PAYE_L3_JUN-2026.xlsx").exists() and (months / "PAYE_L3_JUL-2025.xlsx").exists()


def test_months_inside_the_workbook_count_as_downloaded(tmp_path):
    months = _two_months(tmp_path)
    th.update_tax_head(tmp_path, "PAYE")
    eng.mark_progress(tmp_path, "PAYE", "JUN-2026", "saved")
    assert eng.month_file_done(tmp_path, "PAYE", "JUN-2026", months / "PAYE_L3_JUN-2026.xlsx")
    assert eng.month_file_done(tmp_path, "PAYE", "JUL-2025", months / "PAYE_L3_JUL-2025.xlsx")
    assert not eng.month_file_done(tmp_path, "PAYE", "MAY-2026", months / "PAYE_L3_MAY-2026.xlsx")
    (tmp_path / "PAYE.xlsx").unlink()          # workbook deleted by hand: 'saved' alone is not enough
    assert not eng.month_file_done(tmp_path, "PAYE", "JUN-2026", months / "PAYE_L3_JUN-2026.xlsx")


def test_month_files_kept_when_the_workbook_is_open_in_excel(tmp_path, monkeypatch):
    months = _two_months(tmp_path)
    th.update_tax_head(tmp_path, "PAYE")
    level3(months, "PAYE", "JUL-2026", [row(1, "2270002000001", "1,200,000", "JUL-2026")])
    real = th.os.replace

    def locked(src, dst):
        if str(dst).endswith("PAYE.xlsx"):
            raise PermissionError("open in Excel")
        return real(src, dst)
    monkeypatch.setattr(th.os, "replace", locked)
    with pytest.raises(PermissionError):
        th.update_tax_head(tmp_path, "PAYE")
    assert (months / "PAYE_L3_JUL-2026.xlsx").exists(), "never removed unless it is inside the real workbook"
    monkeypatch.setattr(th.os, "replace", real)
    th.update_tax_head(tmp_path, "PAYE")       # Excel closed: next run folds it in
    assert "JUL-2026" in load_workbook(tmp_path / "PAYE.xlsx").sheetnames and not months.exists()


def test_summary_lists_months_without_level3_and_keeps_them_next_run(tmp_path):
    _two_months(tmp_path)
    th.update_tax_head(tmp_path, "PAYE", statuses={
        "JUN-2026": ("saved", "Payment 1,160,500"), "MAY-2026": ("nil payment - skipped", "Payment 0"),
        "APR-2026": ("not in ledger", "")})
    th.update_tax_head(tmp_path, "PAYE", statuses={"APR-2026": ("not in ledger", "")})   # a later run
    summary = [list(r) for r in load_workbook(tmp_path / "PAYE.xlsx")["Summary"].iter_rows(values_only=True)]
    got = {r[0]: (r[1], r[2]) for r in summary[1:]}
    assert list(got) == ["JUL-2025", "APR-2026", "MAY-2026", "JUN-2026"]
    assert got["JUN-2026"] == ("Level 3 saved", 1160500), "the Level 2 payment as a number"
    assert got["MAY-2026"] == ("nil payment - skipped", 0), "kept from the earlier run"
    assert got["APR-2026"] == ("not in ledger", None)


def test_workbook_from_version_10_9_gets_its_month_tabs(tmp_path):
    """10.9 wrote only the 'Level 3-2' sheet and kept every month in _months: the first 10.12 run
    builds the tabs from those files and then removes them; a month whose file is gone gets a tab
    rebuilt from its table rows."""
    months = _two_months(tmp_path)
    th.update_tax_head(tmp_path, "PAYE", keep_month_files=True)
    wb = load_workbook(tmp_path / "PAYE.xlsx")
    for name in ("Summary", "JUL-2025", "JUN-2026"):
        del wb[name]                                       # = the 10.9 layout
    wb.save(tmp_path / "PAYE.xlsx")
    (months / "PAYE_L3_JUL-2025.xlsx").unlink()            # its file lost
    th.update_tax_head(tmp_path, "PAYE")
    wb = load_workbook(tmp_path / "PAYE.xlsx")
    assert wb.sheetnames == ["Summary", "Level 3-2", "JUL-2025", "JUN-2026"]
    assert "2260003116560" in _cells(wb["JUL-2025"]) and "Sr No." in _cells(wb["JUL-2025"])


def test_route_check_left_in_months_by_10_11_moves_to_the_client_folder(tmp_path):
    import json
    months = th.months_folder(tmp_path, "PAYE")
    months.mkdir(parents=True)
    (months / eng.ROUTE_CHECK_FILE).write_text(json.dumps({"PAYE": {"result": "identical"}}), encoding="utf-8")
    (tmp_path / eng.ROUTE_CHECK_FILE).write_text(json.dumps({"IT-WHT": {"result": "identical"}}), encoding="utf-8")
    eng.move_route_check_up(months)
    assert not (months / eng.ROUTE_CHECK_FILE).exists()
    assert set(json.loads((tmp_path / eng.ROUTE_CHECK_FILE).read_text(encoding="utf-8"))) == {"PAYE", "IT-WHT"}


# ---------- 10.13: Level 2 payment against the month's Level 3-2 lines


def line(n, basis, prn, amount, adjusted, principal=0, penalty=0, interest=0, refund=0):
    return [n, basis, "14-SEP-2026", prn, amount, adjusted, "Original Return", principal, penalty, interest, refund]


def _summary(path):
    rows = list(load_workbook(path)["Summary"].iter_rows(values_only=True))
    head = list(rows[0])
    return {r[0]: dict(zip(head, r)) for r in rows[1:]}


def level2(folder, tax, rows, parts=("PRN", "WHT Credit", "Offset Received(Other Period)")):
    """A Level 2 file as the program saves it: 'Payment' over its sub-columns (VAT layout of the live run)."""
    head = ("<tr><th rowspan='2'>S/N</th><th rowspan='2'>Period</th><th rowspan='2'>Return</th>"
            "<th rowspan='2'>Assessment</th><th rowspan='2'>Penalty</th><th rowspan='2'>Interest</th>"
            f"<th colspan='{len(parts)}'>Payment</th><th rowspan='2'>Net Position</th></tr>"
            "<tr>" + "".join(f"<th>{p}</th>" for p in parts) + "</tr>")
    body = "".join(f"<tr><td>{i}</td><td>{m}</td><td>0</td><td>0</td><td>0</td><td>0</td>"
                   + "".join(f"<td>{v}</td>" for v in vals) + "<td>0</td></tr>" for i, (m, vals) in enumerate(rows, 1))
    folder.mkdir(parents=True, exist_ok=True)
    excel_files.html_to_xlsx('<table><tr><td colspan="9">Tax Position as at 08/10/2026 10:32:30</td></tr>'
                             + head + body + "</table>", folder / f"{tax}_L2.xlsx")


def test_level2_payment_total_is_the_amount_of_each_different_line(tmp_path):
    """10.15 - the rule read from the user's live output of 8 Oct 2026 (the patterns only; every number and
    reference here is invented):
      * one PRN split over two periods is listed twice with the same Amount -> counted ONCE;
      * part of a payment not allocated (Principal less than the Amount) -> Level 2 still counts all of it;
      * a Credit Balance Transfer is in Level 2's 'Offset Received' sub-column - or, for an old one, inside
        the PRN sub-column (VAT AUG-2025 live) -> only the TOTAL of the Payment sub-columns is compared.
    The 10.13 rule (total Principal) called the first two 'check' although Level 2 was right."""
    tax = "Value Added Tax"
    months = th.months_folder(tmp_path, tax)
    level3(months, tax, "FEB-2026", [
        line(1, "Payment (PRN)", "2260000000101", "500,000", "FEB-2026", principal="300,000"),
        line(2, "Payment (PRN)", "2260000000101", "500,000", "JAN-2026", principal="200,000"),
        line(3, "Payment (PRN)", "2260000000102", "40,000", "FEB-2026", principal="40,000")])
    level3(months, tax, "MAY-2026", [
        line(1, "Payment (PRN)", "2260000000103", "900,000", "MAY-2026", principal="650,000"),
        line(2, "Credit Balance Transfer", "MTO000000000011", "75,000", "MAY-2026", principal="75,000")])
    level3(months, tax, "AUG-2025", [
        line(1, "Credit Balance Transfer", "MTO000000000012", "60,000", "AUG-2025", principal="60,000"),
        line(2, "Payment (PRN)", "2260000000105", "1,000,000", "AUG-2025", principal="1,000,000")])
    level3(months, tax, "JUN-2026", [
        line(1, "Payment (PRN)", "2260000000104", "120,000", "JUN-2026", principal="120,000")])
    level2(tmp_path, tax, [("JUN-2026", ("125,000", "0", "0")), ("MAY-2026", ("900,000", "0", "75,000")),
                           ("APR-2026", ("0", "0", "0")), ("FEB-2026", ("540,000", "0", "0")),
                           ("AUG-2025", ("1,060,000", "0", "0")), ("JUL-2025", ("7,000", "0", "0"))])
    checks = []
    th.update_tax_head(tmp_path, tax, checks=checks, statuses={
        "FEB-2026": ("saved", "Payment 540,000"), "MAY-2026": ("saved", "Payment 900,000"),
        "JUN-2026": ("saved", "Payment 125,000"), "AUG-2025": ("saved", "Payment 1,060,000"),
        "APR-2026": ("nil payment - skipped", "Payment 0")})
    got = _summary(tmp_path / f"{tax}.xlsx")
    assert list(got) == ["AUG-2025", "FEB-2026", "APR-2026", "MAY-2026", "JUN-2026"], \
        "Level 2 periods this run did not touch are not added"
    assert (got["FEB-2026"]["Level 3 total"], got["FEB-2026"]["Check"]) == (540000, "OK")
    assert got["FEB-2026"]["Not allocated"] is None, "500,000 fully allocated over its two lines"
    may = got["MAY-2026"]
    assert (may["Payment (Level 2)"], may["Level 3 total"], may["Check"]) == (975000, 975000, "OK")
    assert may["Level 2 split"] == "PRN 900,000; Offset Received(Other Period) 75,000"
    assert (may["Level 3 PRN payments"], may["Level 3 credits"]) == (900000, 75000)
    assert (may["Credits by type"], may["Not allocated"]) == ("Credit Balance Transfer 75,000", 250000)
    assert got["AUG-2025"]["Check"] == "OK", "old transfer inside Level 2's PRN column: the totals agree"
    assert got["JUN-2026"]["Difference"] == 5000
    assert got["JUN-2026"]["Check"] == ("check: Level 2 payment 125,000, Level 3 lines 120,000 (PRN payments "
                                        "120,000, 1 PRN(s)), differs by 5,000")
    assert (got["APR-2026"]["Difference"], got["APR-2026"]["Check"]) == (None, None), "nil month: nothing to compare"
    assert [p for p, _ in checks] == ["JUN-2026"], "only a real difference is a check"


def test_check_names_other_periods_a_line_was_adjusted_against(tmp_path):
    months = th.months_folder(tmp_path, "IT-WHT")
    level3(months, "IT-WHT", "JUL-2025", [line(1, "Payment (PRN)", "2260000000003", "172,456", "JUL-2025"),
                                          line(2, "Payment (PRN)", "2260000000004", "26,745", "AUG-2025")])
    checks = []
    th.update_tax_head(tmp_path, "IT-WHT", checks=checks, statuses={"JUL-2025": ("saved", "Payment 172,456")})
    got = _summary(tmp_path / "IT-WHT.xlsx")["JUL-2025"]
    assert got["Difference"] == -26745
    assert "differs by -26,745" in got["Check"] and "adjusted against AUG-2025" in got["Check"]


def test_yearly_tax_head_checks_payment_and_tax_credit(tmp_path):
    """Income Tax capture of 8 Oct 2026 (structure only): Level 2 of a year has PRN and Tax Credit columns,
    both opening the same Level 3-2 page; Tax Credit = the WHT Credit + ASYCUDA Credit lines."""
    year = "JUL-2026 to JUN-2027"
    months = th.months_folder(tmp_path, "Income Tax")
    level3(months, "Income Tax", year, [
        line(1, "Payment (PRN)", "2260000000201", "100,000", "NA"),
        line(2, "Payment (PRN)", "2260000000202", "250,000", "NA"),
        line(3, "WHT Credit", "MTO000000000001", "60,000", "NA"),
        line(4, "ASYCUDA Credit", "C 0000 2026", "400,000", "NA")])
    th.update_tax_head(tmp_path, "Income Tax", statuses={year: ("saved", "Payment 350,000; Tax credit 460,000"),
                                                        "JUL-2025 to JUN-2026": ("no payment link",
                                                                                 "Payment 0; Tax credit 0")})
    wb = load_workbook(tmp_path / "Income Tax.xlsx")
    assert wb.sheetnames == ["Summary", "Level 3-2", year], "one tab per year"
    got = _summary(tmp_path / "Income Tax.xlsx")
    assert list(got) == ["JUL-2025 to JUN-2026", year], "oldest year first"
    y = got[year]
    assert (y["Payment (Level 2)"], y["Level 2 split"]) == (810000, "PRN 350,000; Tax Credit 460,000")
    assert (y["Level 3 total"], y["Level 3 PRN payments"], y["Level 3 credits"]) == (810000, 350000, 460000)
    assert y["Credits by type"] == "WHT Credit 60,000; ASYCUDA Credit 400,000"
    assert (y["Difference"], y["Check"]) == (0, "OK")
    assert th.months_in_workbook(tmp_path, "Income Tax") == {year}
    th.update_tax_head(tmp_path, "Income Tax", statuses={year: ("saved", "Payment 350,000; Tax credit 470,000")})
    y = _summary(tmp_path / "Income Tax.xlsx")[year]
    assert y["Check"] == ("check: Level 2 payment 820,000, Level 3 lines 810,000 (PRN payments 350,000, "
                          "2 PRN(s); credits 460,000), differs by 10,000")
