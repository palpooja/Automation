"""One Ledger month, Level 2 -> Level 3 -> Export -> Back (6 Oct 2026):
* timing line per month, so the next live log shows where the time goes (Level 3 opening, the file,
  getting back to Level 2) - the old log had only 'opening payment' and 'Saved';
* option B check: is the on-screen Level 3, put through the SAME converter as the portal's file,
  identical to the portal's file?  The saved file is always the portal's; the check only reports;
* no fixed pop-up wait after the month (universal rule).
"""
import json
import re

import pytest
from openpyxl import load_workbook

from engines import ledger_engine as eng

pytestmark = pytest.mark.browser

TIMES = re.compile(r"JUN-2026: times - Level 3 opened after (\d+) s, file saved (\d+) s later "
                   r"\(portal export\), back on Level 2 (\d+) s later")


def _month(page, mock_url, tmp_path, mode):
    page.goto(f"{mock_url}/Portal/ledger_l2l3.html")
    page.evaluate("m => sessionStorage.setItem('exportMode', m)", mode)
    eng.TIMINGS.clear()
    return eng.download_level3_for_month(page, page.main_frame, tmp_path, "PAYE", "JUN-2026", "12,487,927")


def test_month_logs_its_times_and_an_identical_screen_copy(page, mock_url, tmp_path, capsys, monkeypatch):
    monkeypatch.setattr(eng, "LEVEL3_SCREEN_CHECK", True)
    _, outcome = _month(page, mock_url, tmp_path, "same")
    assert outcome == "saved"
    out = capsys.readouterr().out
    m = TIMES.search(out)
    assert m, out
    assert int(m.group(3)) < 5, "no fixed 5 s pop-up wait between the file and Level 2"
    assert "JUN-2026: screen copy identical to the portal's file" in out, out
    checks = json.loads((tmp_path / eng.SCREEN_CHECK_FILE).read_text(encoding="utf-8"))
    assert checks["PAYE_L3_JUN-2026.xlsx"]["result"] == "identical"
    assert eng.TIMINGS and eng.TIMINGS[0][:2] == ("PAYE", "JUN-2026")


def test_a_different_screen_copy_is_reported_and_the_portal_file_is_kept(page, mock_url, tmp_path, capsys,
                                                                      monkeypatch):
    monkeypatch.setattr(eng, "LEVEL3_SCREEN_CHECK", True)
    _, outcome = _month(page, mock_url, tmp_path, "extra")
    assert outcome == "saved"
    out = capsys.readouterr().out
    assert "JUN-2026: screen copy DIFFERENT from the portal's file" in out, out
    checks = json.loads((tmp_path / eng.SCREEN_CHECK_FILE).read_text(encoding="utf-8"))
    assert checks["PAYE_L3_JUN-2026.xlsx"]["result"] == "different"
    saved = load_workbook(tmp_path / "PAYE_L3_JUN-2026.xlsx")
    assert saved.worksheets[0]["A1"].value == "URA export heading"     # still the portal's own file
    assert not list(tmp_path.glob("*screen*.xlsx")), "no scratch workbook left beside the downloads"


def test_summary_of_times_and_screen_checks(tmp_path, capsys):
    eng.TIMINGS[:] = [("PAYE", "JUN-2026", 20, 10, 40), ("PAYE", "MAY-2026", 30, 12, 50),
                      ("VAT", "JUN-2026", 22, 9, 44)]
    (tmp_path / eng.SCREEN_CHECK_FILE).write_text(json.dumps({
        "PAYE_L3_JUN-2026.xlsx": {"result": "identical"}, "PAYE_L3_MAY-2026.xlsx": {"result": "identical"},
        "VAT_L3_JUN-2026.xlsx": {"result": "different"}}), encoding="utf-8")
    eng.summarise_level3_checks(tmp_path)
    out = capsys.readouterr().out
    assert "Level 3 months timed: 3 - median: Level 3 opened after 22 s, file saved 10 s later, " \
           "back on Level 2 44 s later" in out, out
    assert "Screen-copy check: 2 of 3 Level 3 file(s) identical, 1 different" in out, out


def test_screen_copy_check_is_off_by_default_and_costs_nothing(page, mock_url, tmp_path, capsys):
    """Live run 6 Oct 2026: 0 of 27 identical; option C made the portal's file take 1 s."""
    assert eng.LEVEL3_SCREEN_CHECK is False
    page.goto(f"{mock_url}/Portal/ledger_l2l3.html")
    eng.download_level3_for_month(page, page.main_frame, tmp_path, "PAYE", "JUN-2026", "12,487,927")
    assert "screen copy" not in capsys.readouterr().out
    assert not (tmp_path / eng.SCREEN_CHECK_FILE).exists()
