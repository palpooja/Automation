"""Behaviour seen in the live run of 3 Oct 2026 (log run_20261003_235216):
1. Logout asks 'You have requested to be logged out, do you want to proceed?' - it must be answered OK,
   and 'Logged out' may only be reported when the portal really logged out.
2. The calendar ignores mouse clicks; each arrow and day click cost 4 s before the script click worked."""
import time

import pytest

from core import session
from engines import wht_engine as we, wht_outputs as wo

pytestmark = pytest.mark.browser


def _logged_out(page):
    return all(f.locator("#logoutFormISSP").count() == 0 for f in page.frames)


@pytest.mark.parametrize("certificates_ran", [False, True], ids=["returns-ledgers-only", "after-certificates"])
def test_logout_answers_the_logout_question_and_really_logs_out(page, mock_url, certificates_ran):
    page.goto(f"{mock_url}/home.html")
    if certificates_ran:              # the WHT engine's question policy is attached to this page
        we.WhtBot(page, lambda *_: None, lambda: None)
    logs = []
    assert session.logout(page, lambda m, *a: logs.append(m)) is True
    assert _logged_out(page)
    assert any("Logged out of the URA portal" in m for m in logs)


def test_logout_says_so_when_the_portal_did_not_log_out(page):
    page.set_content('<form id="logoutFormISSP"><input type="button" value="Logout" onclick="1"></form>')
    logs = []
    assert session.logout(page, lambda m, *a: logs.append(m)) is False
    assert not any(m.startswith("Logged out of the URA portal") for m in logs)


def test_an_unrelated_question_during_logout_is_still_answered_cancel(page):
    page.set_content('<script>window.r = null;</script><form id="logoutFormISSP"><input type="button" '
                     'value="Logout" onclick="window.r = confirm(\'Delete all certificates?\')"></form>')
    session.logout(page, lambda *a: None)
    assert page.evaluate("() => window.r") is False


def test_calendar_that_ignores_mouse_clicks_costs_no_extra_time(page, mock_url, monkeypatch):
    monkeypatch.setattr(we, "RESULT_WAIT_MS", 30000)
    page.goto(f"{mock_url}/Portal/wht_slowclick.html")
    bot = we.WhtBot(page, lambda *_: None, lambda: None)
    frame = bot.form_frame()
    t0 = time.monotonic()
    warnings = bot.pick_month(frame, 2026, 6)          # mock calendar opens on October 2026: 4 steps back
    took = time.monotonic() - t0
    assert warnings == []
    assert frame.locator("#frmDt").input_value() == "01/06/2026"
    assert frame.locator("#toDt").input_value() == "30/06/2026"
    assert took < 6, f"took {took:.1f} s (the old code needed about 4 s per arrow and 4 s for day 1)"


def test_calendar_that_ignores_mouse_clicks_gives_the_same_results(page, mock_url, monkeypatch):
    monkeypatch.setattr(we, "RESULT_WAIT_MS", 30000)
    page.goto(f"{mock_url}/Portal/wht_slowclick.html")
    bot = we.WhtBot(page, lambda *_: None, lambda: None)
    months = [wo.MonthResult(2026, m) for m in (1, 2, 3)]
    we.run_months(bot, months, want_pdf=False, want_html=False, generate_for="D")
    assert [(r.label, r.status, r.rows, r.warnings) for r in months] == [
        ("2026-01", "empty", 0, []), ("2026-02", "ok", 2, []), ("2026-03", "ok", 1, [])]
