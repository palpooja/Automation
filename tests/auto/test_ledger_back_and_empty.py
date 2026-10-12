"""10.17 - two costs seen in both live runs of 9 Oct 2026 (Tronics, Telecom):
* 'still on Level 3 - pressing Back again' on the first month of every tax type (11-29 s each): the Back had
  worked, Level 2 was only slow to come back, and the program pressed again after a fixed 10 s;
* an empty Local Excise Duty ledger took 21-25 s to save: the Level 2 table was waited for by a month row that a
  ledger with no period never shows.
Test portal: tests/mock_portal/ledger_server.py (invented values)."""
import time

import pytest

from engines import ledger_engine as eng

pytestmark = pytest.mark.browser


def _level2(page, mock_url, ledger_portal, **state):
    ledger_portal.reset(**state)
    page.goto(f"{mock_url}{ledger_portal.BASE}loadLevelTwo")
    eng.TIMINGS.clear()
    return page.main_frame


def test_a_back_that_worked_is_not_pressed_again_while_level2_is_slow(page, mock_url, ledger_portal, tmp_path,
                                                                    capsys):
    frame = _level2(page, mock_url, ledger_portal, level2_delay=12)      # Level 2 takes 12 s to come back
    _, outcome = eng.download_level3_for_month(page, frame, tmp_path, "PAYE", "JUN-2026", "1,329,127")
    assert outcome == "saved"
    out = capsys.readouterr().out
    assert "pressing Back again" not in out, out
    assert ledger_portal.STATE["level2_loads"] == 2, "the first load and ONE Back - not a second Back"


def test_a_back_press_that_did_nothing_is_pressed_again_after_a_few_seconds(page, mock_url, ledger_portal, tmp_path,
                                                                           capsys, monkeypatch):
    monkeypatch.setattr(eng, "BACK_TOOK_S", 3)
    frame = _level2(page, mock_url, ledger_portal, dead_backs=1)        # the first Back press is ignored
    t0 = time.monotonic()
    _, outcome = eng.download_level3_for_month(page, frame, tmp_path, "PAYE", "JUN-2026", "1,329,127")
    took = time.monotonic() - t0
    assert outcome == "saved"
    out = capsys.readouterr().out
    assert "asked the portal for nothing in 3 s - pressing Back again (1/3)" in out, out
    assert ledger_portal.STATE["level2_loads"] == 2
    assert took < 25, f"{took:.0f} s"


def test_an_empty_ledger_is_read_and_saved_within_seconds(page, mock_url, ledger_portal, tmp_path):
    frame = _level2(page, mock_url, ledger_portal, no_months=True)
    t0 = time.monotonic()
    df = eng.read_level2_table(frame)
    assert len(df) == 0 and "Period" in df.columns and "Payment" in df.columns
    how = eng.export_level2_to_excel(page, frame, tmp_path, "Local Excise Duty", has_months=False)
    took = time.monotonic() - t0
    assert how == "portal export", how
    assert (tmp_path / "Local Excise Duty_L2.xlsx").exists()
    assert took < 8, f"{took:.0f} s (was 20 s waiting for a month row)"


def test_a_ledger_with_months_still_waits_for_its_rows(page, mock_url, ledger_portal, tmp_path):
    frame = _level2(page, mock_url, ledger_portal)
    how = eng.export_level2_to_excel(page, frame, tmp_path, "PAYE")
    assert how == "portal export"
    assert (tmp_path / "PAYE_L2.xlsx").exists()
