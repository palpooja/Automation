"""Universal rule (user, 6 Oct 2026): a pop-up or notice that may not appear is checked at once and
closed only if it is showing - never waited for.  Measured on 10.6: with nothing on screen,
dismiss_pre_login_popup waited 5.00 s (it runs at login, at the Tax Statement and after every
Level 3 month).  In the 5 Oct logs the pop-up was closed 10 times and every time it was already
showing (closed 0-2 s after the previous step), so checking at once closes the same pop-ups.

Waits for things the program NEEDS (a Level 3 page, a download, the schedules button) are not
covered: checking those at once could lose data.
"""
import time

import pytest

from engines import ledger_engine as eng, returns_pw, wht_engine as we

pytestmark = pytest.mark.browser

POPUP = """<div id="process-static-modal" style="display:%s;position:fixed;top:0;left:0;right:0;
  bottom:0;background:#0006"><div style="background:#fff;margin:60px;padding:20px">Notice
  <button onclick="document.getElementById('process-static-modal').style.display='none'">Close</button>
  </div></div><p>portal page</p>"""
NOTHING = "<p>portal page</p><iframe srcdoc='<p>inner frame</p>'></iframe>"


def _took(fn):
    t0 = time.monotonic()
    fn()
    return time.monotonic() - t0


def test_absent_pre_login_popup_costs_no_time(page):
    page.set_content(NOTHING)
    assert _took(lambda: eng.dismiss_pre_login_popup(page)) < 0.5


def test_hidden_pre_login_popup_is_not_waited_for(page):
    page.set_content(POPUP % "none")
    assert _took(lambda: eng.dismiss_pre_login_popup(page)) < 0.5


def test_showing_pre_login_popup_is_still_closed(page, capsys):
    page.set_content(POPUP % "block")
    eng.dismiss_pre_login_popup(page)
    assert not page.locator("#process-static-modal").is_visible()
    assert "Dismissed pre-login popup." in capsys.readouterr().out


ROUTINES = {
    "ledger dismiss_cookie_banner": lambda pg: eng.dismiss_cookie_banner(pg),
    "ledger dismiss_pre_login_popup": lambda pg: eng.dismiss_pre_login_popup(pg),
    "ledger clear_overlays": lambda pg: eng.clear_overlays(pg),
    "ledger dismiss_popup_if_present": lambda pg: eng.dismiss_popup_if_present(pg),
    "wht dismiss_popups": lambda pg: we.WhtBot(pg, lambda *a: None, lambda: None).dismiss_popups(),
    "wht _quick_dismiss": lambda pg: we.WhtBot(pg, lambda *a: None, lambda: None)._quick_dismiss(),
    "returns _skip_popups": lambda pg: returns_pw._skip_popups(pg, lambda *a: None),
}


@pytest.mark.parametrize("name", sorted(ROUTINES))
def test_every_popup_routine_returns_at_once_when_nothing_is_showing(page, name):
    page.set_content(NOTHING)
    took = _took(lambda: ROUTINES[name](page))
    assert took < 1.0, f"{name} took {took:.2f} s with no pop-up on screen"


def test_logout_link_is_pressed_without_waiting_on_a_hidden_one(page):
    from core import session
    page.set_content("""<a href="#" style="display:none" onclick="window.hidden=1">Logout</a>
        <form id="logoutFormISSP"><a href="#" onclick="window.hit=1;return false;">Logout</a></form>""")
    took = _took(lambda: session.press_logout_link_if_shown(page))
    assert page.evaluate("() => window.hit") == 1 and page.evaluate("() => window.hidden") is None
    assert took < 1.0, f"took {took:.2f} s"
