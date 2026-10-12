"""WHT certificates, live run of 5 Oct 2026 (run_20261005_171158, screenshots screen_error_*.png).

The URA page uses 'The DHTML Calendar' (tests/mock_portal/Portal/jscalendar/SOURCE.txt).  Its
buttons react to a mouse button pressed and released on them, not to a plain 'click' signal, so:
  1. the program's script clicks never moved the month ("Stepped 60 month(s) to 'October, 2026'");
  2. the calendar stayed open over the Search button;
  3. the forced Search click landed on the calendar and replaced the From date with a Tuesday;
  4. Search was never pressed and the month was reported as 'portal too slow'.
These tests use the real calendar code on a copy of the page laid out like the portal.
"""
import time

import pytest

from engines import wht_engine as we, wht_outputs as wo

pytestmark = pytest.mark.browser

PAGE = "/Portal/wht_dynarch.html"

# The calendar's month arrows ignore every press (an extreme portal fault), used to prove that the
# program gives up quickly and still never searches with a calendar lying over the Search button.
DEAD_ARROWS_JS = """() => { const orig = Calendar.cellClick;
  Calendar.cellClick = function (el, ev) { if (el.navtype && Math.abs(el.navtype) <= 2) return;
                                           return orig(el, ev); }; }"""


def _open(page, mock_url, from_value):
    """The page as the portal shows it after a reload: the From box holds the last date sent."""
    page.goto(f"{mock_url}{PAGE}")
    page.evaluate("() => sessionStorage.clear()")
    page.evaluate("v => { frmDt.value = v; }", from_value)
    logs = []
    bot = we.WhtBot(page, lambda m, *a: logs.append(m), lambda: None)
    return bot, bot.form_frame(), logs


def _calendar_open(frame):
    return frame.evaluate("""() => [...document.querySelectorAll('.jscalendar')]
        .some(d => d.style.display !== 'none' && d.getBoundingClientRect().height > 0)""")


def _searches(page):
    return page.evaluate("() => JSON.parse(sessionStorage.getItem('searches') || '[]')")


def test_real_calendar_moves_to_the_month_and_picks_day_1(page, mock_url):
    bot, frame, logs = _open(page, mock_url, "05/10/2026")       # opens on October 2026
    t0 = time.monotonic()
    warnings = bot.pick_month(frame, 2025, 7)
    took = time.monotonic() - t0
    assert warnings == [], logs
    assert frame.locator("#frmDt").input_value() == "01/07/2025"
    assert frame.locator("#toDt").input_value() == "31/07/2025"
    assert not _calendar_open(frame), "the calendar must be closed so it cannot cover Search"
    assert any("Period set to 01/07/2025" in m for m in logs), logs
    assert not any("written directly" in m for m in logs), logs
    # the log must count months really moved (October 2026 -> July 2025 = 15), not presses
    assert any("Moved 15 month(s)" in m for m in logs), logs
    assert took < 20, f"took {took:.1f} s"


def test_search_is_really_sent_with_the_chosen_month(page, mock_url, monkeypatch):
    monkeypatch.setattr(we, "RESULT_WAIT_MS", 15000)
    bot, frame, logs = _open(page, mock_url, "05/07/2025")       # opens on July 2025: no walking
    months = [wo.MonthResult(2025, 7)]
    we.run_months(bot, months, want_pdf=False, want_html=False, generate_for="D")
    assert _searches(page) == [["01/07/2025", "31/07/2025", "D"]], logs
    assert (months[0].status, months[0].rows) == ("ok", 1), (months[0].status, months[0].message, logs)
    assert not any("portal too slow" in m for m in logs), logs


def test_two_months_in_a_row_each_search_their_own_dates(page, mock_url, monkeypatch):
    monkeypatch.setattr(we, "RESULT_WAIT_MS", 15000)
    bot, frame, logs = _open(page, mock_url, "05/10/2026")
    months = [wo.MonthResult(2025, 7), wo.MonthResult(2025, 8)]
    we.run_months(bot, months, want_pdf=False, want_html=False, generate_for="D")
    assert _searches(page) == [["01/07/2025", "31/07/2025", "D"], ["01/08/2025", "31/08/2025", "D"]], logs
    assert [(r.status, r.rows, r.warnings) for r in months] == [("ok", 1, []), ("ok", 2, [])], logs


def test_search_never_presses_the_screen_where_the_calendar_lies(page, mock_url, monkeypatch):
    """Exactly the 5 Oct state: From correct, calendar left open over the Search button."""
    monkeypatch.setattr(we, "RESULT_WAIT_MS", 10000)
    bot, frame, logs = _open(page, mock_url, "01/07/2025")
    frame.evaluate("() => getToDate(frmDt)")
    frame.locator("#frmDtImg").evaluate("el => el.click()")          # calendar opens and covers Search
    assert _calendar_open(frame)
    bot.expected_from = "01/07/2025"
    assert bot.search(frame) == "table", logs
    assert _searches(page) == [["01/07/2025", "31/07/2025", ""]], logs


def test_calendar_whose_arrows_do_not_move_is_given_up_quickly_and_closed(page, mock_url, monkeypatch):
    monkeypatch.setattr(we, "RESULT_WAIT_MS", 10000)
    bot, frame, logs = _open(page, mock_url, "05/10/2026")
    frame.evaluate(DEAD_ARROWS_JS)
    t0 = time.monotonic()
    warnings = bot.pick_month(frame, 2025, 7)
    took = time.monotonic() - t0
    assert took < 15, f"took {took:.1f} s (5 Oct: about 110 s per month pressing a dead arrow)"
    assert any("did not move" in m for m in logs), logs
    assert any("written directly" in w for w in warnings), warnings
    assert frame.locator("#frmDt").input_value() == "01/07/2025"
    assert frame.locator("#toDt").input_value() == "31/07/2025"
    assert not _calendar_open(frame), "a calendar left open is what changed the From date on 5 Oct"
    bot.search(frame)
    assert _searches(page)[-1][:2] == ["01/07/2025", "31/07/2025"], logs


def test_wrong_from_date_is_reported_as_such_not_as_a_slow_portal(page, mock_url, monkeypatch):
    monkeypatch.setattr(we, "RESULT_WAIT_MS", 10000)
    bot, frame, logs = _open(page, mock_url, "22/07/2025")
    frame.evaluate("() => getToDate(frmDt)")
    bot.expected_from = "01/10/2025"
    with pytest.raises(RuntimeError) as err:
        bot.search(frame)
    assert "22/07/2025" in str(err.value) and "01/10/2025" in str(err.value)
    assert "too slow" not in str(err.value)
    assert _searches(page) == [], "Search must not be pressed with a wrong From date"


def test_january_week_number_1_is_not_taken_for_day_1(page, mock_url):
    """In a January view the calendar's 'wk' column shows week 1 before the real day 1."""
    bot, frame, logs = _open(page, mock_url, "05/02/2026")
    warnings = bot.pick_month(frame, 2026, 1)
    assert warnings == [], logs
    assert frame.locator("#frmDt").input_value() == "01/01/2026"
    assert frame.locator("#toDt").input_value() == "31/01/2026"
    assert not any("written directly" in m for m in logs), logs
