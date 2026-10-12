"""Option C (6 Oct 2026): each month's Level 3 loaded into a hidden frame, so Level 2 stays on screen
and no Back (= the portal rebuilding Level 2, 46-68 s per month in the 5 Oct logs) is needed.

Built from the user's capture: the Payment link is <a onclick="openLevelThreeTwo('AUG-2026')">, and
openLevelThreeTwo() points the POST form 'levelTwoOfAR' at loadLevelThreeTwo?period=... and submits it.
The program calls the link's own onclick with the form's target set to a hidden frame - the portal
receives exactly the request a click sends.  Test portal: tests/mock_portal/ledger_server.py.
"""
import json

import pytest

from engines import excel_files, ledger_engine as eng

pytestmark = pytest.mark.browser
MONTHS = [("JUN-2026", "1,329,127"), ("MAY-2026", "8,068,063"), ("APR-2026", "9,704,197")]
TIME_LINE = ("Tax Position as at",)


@pytest.fixture
def ledger(page, mock_url, ledger_portal, monkeypatch):
    monkeypatch.setattr(eng, "L3_ROUTE", "screen")        # these tests are about the screen way (10.16: not the default)
    eng.DIRECT_STATE.clear()

    def start(**state):
        ledger_portal.reset(**state)
        page.goto(f"{mock_url}{ledger_portal.BASE}loadLevelTwo")
        eng.TIMINGS.clear()
        eng.ROUTE_STATE.clear()
        return page.main_frame
    # a Level 2 freshly loaded from the portal (the real code goes Level 1 -> Level 2)
    monkeypatch.setattr(eng, "refresh_level2",
                        lambda pg, fr, tax: (pg.goto(f"{mock_url}{ledger_portal.BASE}loadLevelTwo"), pg.main_frame)[1])
    return start


def _months(page, frame, folder, months=MONTHS):
    out = []
    for m, amt in months:
        frame, outcome = eng.level3_month(page, frame, folder, "PAYE", m, amt)
        out.append(outcome)
    return out


def test_background_route_saves_months_without_rebuilding_level2(page, ledger, ledger_portal, tmp_path,
                                                                  monkeypatch, capsys):
    monkeypatch.setattr(eng, "HIDDEN_L3_VERIFY", False)
    frame = ledger()
    loads = ledger_portal.STATE["level2_loads"]
    assert _months(page, frame, tmp_path) == ["saved"] * 3
    assert ledger_portal.STATE["level2_loads"] == loads, "no Back: Level 2 must never be rebuilt"
    assert all((tmp_path / f"PAYE_L3_{m}.xlsx").exists() for m, _ in MONTHS)
    assert capsys.readouterr().out.count("Level 2 kept (no Back)") == 3


def test_background_files_equal_click_route_files_except_the_time_line(page, ledger, ledger_portal,
                                                                      tmp_path, monkeypatch):
    monkeypatch.setattr(eng, "HIDDEN_L3_VERIFY", False)
    hidden, click = tmp_path / "hidden", tmp_path / "click"
    _months(page, ledger(), hidden)
    monkeypatch.setattr(eng, "HIDDEN_L3", False)
    _months(page, ledger(), click)
    for m, _ in MONTHS:
        a, b = hidden / f"PAYE_L3_{m}.xlsx", click / f"PAYE_L3_{m}.xlsx"
        assert excel_files.compare_cells(a, b, ignore_text=TIME_LINE) == (True, [])
        assert excel_files.compare_cells(a, b)[0] is False, "the time line differs on every download"
        assert sorted(p.name for p in (hidden / "portal_original").iterdir()) == \
            sorted(p.name for p in (click / "portal_original").iterdir())


def test_first_month_is_proved_against_the_click_route_then_background(page, ledger, ledger_portal,
                                                                      tmp_path, capsys):
    frame = ledger()
    loads = ledger_portal.STATE["level2_loads"]
    assert _months(page, frame, tmp_path) == ["saved"] * 3
    assert ledger_portal.STATE["level3_requests"] == ["JUN-2026", "JUN-2026", "MAY-2026", "APR-2026"]
    assert ledger_portal.STATE["level2_loads"] == loads + 1, "only the first month uses Back"
    check = json.loads((tmp_path / eng.ROUTE_CHECK_FILE).read_text(encoding="utf-8"))
    assert check["PAYE"]["result"] == "identical", check
    out = capsys.readouterr().out
    assert "background copy identical" in out and out.count("Level 2 kept (no Back)") == 2


def test_portal_refusing_a_repeated_form_falls_back_to_click_and_back(page, ledger, ledger_portal,
                                                                       tmp_path, capsys):
    frame = ledger(csrf="rotate")          # every token usable once - the strictest portal
    assert _months(page, frame, tmp_path) == ["saved"] * 3
    out = capsys.readouterr().out
    assert "usual click-and-Back way for the rest of PAYE" in out, out
    for m, _ in MONTHS:
        assert (tmp_path / f"PAYE_L3_{m}.xlsx").exists()


def test_summary_separates_the_two_routes(tmp_path, capsys):
    eng.TIMINGS[:] = [("PAYE", "JUN-2026", 20, 10, 40, "click"), ("PAYE", "MAY-2026", 22, 9, 0, "background"),
                      ("PAYE", "APR-2026", 24, 11, 0, "background")]
    eng.summarise_level3_checks(tmp_path)
    out = capsys.readouterr().out
    assert "Level 3 months timed: 1 - median: Level 3 opened after 20 s, file saved 10 s later, " \
           "back on Level 2 40 s later." in out, out
    assert "Level 3 months in the background: 2 - median: Level 3 opened after 23 s, file saved 10 s later " \
           "(no Back)." in out, out
