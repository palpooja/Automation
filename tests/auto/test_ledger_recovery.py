"""Ledgers, live runs of 5 Oct 2026 (run_20261005_162406, run_20261005_171158, client B 18:11).

1. A 45-s wait for the Level 3-2 box in a frame that did not hold it, while the screenshots show
   the box on screen (<TIN>_error_PAYE_1.png, <TIN>_error_Value_Added_Tax_2.png).
2. Level 1 given only 10 s to show (it was there 1 s after each failure).
3. 'Getting the menu back (Portal link)' never clicked anything (5 times, no 'Portal link: clicking'
   line), then waited 90 s for a 'Reports' heading that could not appear.
4. That failure escaped the retry loop and ended the whole Ledgers step: attempts 2-3 of PAYE and
   the remaining tax types never ran.
"""
import pytest

from engines import ledger_engine as eng


def _settings(tmp_path):
    return {"tin": "1000000001", "password": "x", "period_from": "2025-07", "period_to": "2026-06",
            "tax_types": ["Value Added Tax", "PAYE"], "other_jobs": []}


@pytest.fixture
def quiet_job(monkeypatch, tmp_path):
    """run_job_on_page without a portal: everything that touches the browser is replaced."""
    for name in ("open_tax_statement", "debug_dump", "dismiss_popup_if_present", "safe_screenshot",
                 "build_combined_workbook", "write_summary"):
        monkeypatch.setattr(eng, name, lambda *a, **k: None)
    monkeypatch.setattr(eng, "OUTPUT_ROOT", str(tmp_path))
    monkeypatch.setattr(eng, "ONLINE_WAIT", None)
    return monkeypatch


def test_a_failed_recovery_does_not_end_the_ledgers_step(quiet_job, tmp_path):
    calls = []

    def process(page, tax, *a, **k):
        calls.append(tax)
        if tax == "Value Added Tax":
            raise RuntimeError("Level 3 box not found")

    def recover(*a, **k):
        raise TimeoutError("Could not find a frame containing heading 'Reports' within 90.0s")

    quiet_job.setattr(eng, "process_tax_type", process)
    quiet_job.setattr(eng, "recover_to_level1", recover)
    failed = eng.run_job_on_page(object(), _settings(tmp_path))
    assert calls == ["Value Added Tax"] * 3 + ["PAYE"], "all 3 attempts, then the next tax type"
    assert failed == ["Value Added Tax"]


def test_a_failed_return_to_level1_after_a_finished_tax_type_is_not_a_failed_tax_type(
        quiet_job, tmp_path, capsys):
    """run 17:11: VAT finished, the return to Level 1 failed, and all of VAT was attempted again."""
    import pandas as pd

    def recover(*a, **k):
        raise TimeoutError("Could not find a frame containing heading 'Reports' within 90.0s")

    quiet_job.setattr(eng, "recover_to_level1", recover)
    quiet_job.setattr(eng, "open_level2", lambda *a, **k: object())
    quiet_job.setattr(eng, "read_level2_table", lambda f: pd.DataFrame({"Period": [], "Payment": []}))
    quiet_job.setattr(eng, "level2_usable", lambda p: False)
    quiet_job.setattr(eng, "export_level2_to_excel", lambda *a, **k: "portal export")
    quiet_job.setattr(eng, "generate_month_labels", lambda a, b: [])
    eng.process_tax_type(object(), "Value Added Tax", "2025-07", "2026-06", tmp_path, "1", "x")
    out = capsys.readouterr().out
    assert "=== Value Added Tax: finished ===" in out
    assert "could not return to level 1" in out.lower()


@pytest.mark.browser
def test_portal_link_is_clicked_even_when_a_hidden_portal_entry_comes_first(page, mock_url, capsys):
    page.goto(f"{mock_url}/Portal/portal_links.html")
    assert eng.click_first_visible(page, ["Portal"], "Portal link") is True
    assert page.evaluate("() => window.hit") == 1
    assert page.evaluate("() => window.hidden_hit") is None
    assert "Portal link: clicking" in capsys.readouterr().out


@pytest.mark.browser
def test_slow_level1_is_waited_for_instead_of_leaving_for_the_menu(page, mock_url, monkeypatch, capsys):
    monkeypatch.setattr(eng, "SLOW_FACTOR", 0.25)     # scaled: old 10 s -> 2.5 s, page needs 4 s
    page.goto(f"{mock_url}/Portal/level1_late.html")
    eng.recover_to_level1(page, None, None)
    out = capsys.readouterr().out
    assert "Back on Level 1." in out
    assert "Getting the menu back" not in out


@pytest.mark.browser
def test_level3_box_is_read_from_the_frame_that_shows_it(page, mock_url):
    page.goto(f"{mock_url}/Portal/level3_outer.html")
    # the mechanism: by its address the outer frame counts as 'Level 3', but it never holds the box
    outer = eng.wait_for_level(page, "Level 3", timeout_seconds=5)
    assert outer is not None and outer.locator("#levelThreeTwoOfAR").count() == 0
    frame, text = eng.find_level3_box(page, "JUN-2026", timeout_seconds=10)
    assert frame.locator("#levelThreeTwoOfAR").count() == 1
    assert "JUN-2026" in text
