"""10.17: the Logout button is on the portal home page.  A client whose last step ended on another screen (live
9 Oct 2026: the PAYE certificate form) got 'Could not find the Logout button' and the URA session stayed open.
Now the portal home page is opened first, once, and Logout pressed there."""
import pytest

from core import session
from engines import ledger_engine as eng

pytestmark = pytest.mark.browser

HOME = """<html><body><div id="mailbox"><a href="#">Reports</a></div>
<form id="logoutFormISSP"><input type="button" value="Logout" onclick="document.body.innerHTML='<p>Bye</p>'"></form>
</body></html>"""


def test_logout_opens_the_home_page_when_the_button_is_not_on_the_current_screen(page, monkeypatch):
    page.set_content("<html><body><form id='TaxCreditCertForm'><select id='period'></select></form></body></html>")
    went_home = []

    def home(pg, *a, **k):
        went_home.append(1)
        pg.set_content(HOME)
        return pg.main_frame
    monkeypatch.setattr(eng, "ensure_portal_home", home)
    logs = []
    assert session.logout(page, lambda m, *a: logs.append(m)) is True
    assert went_home == [1]
    assert any("opening the portal home page first" in m for m in logs), logs
    assert logs[-1] == "Logged out of the URA portal."


def test_logout_on_the_home_page_does_not_go_anywhere_first(page, monkeypatch):
    page.set_content(HOME)
    monkeypatch.setattr(eng, "ensure_portal_home", lambda *a, **k: (_ for _ in ()).throw(AssertionError("not needed")))
    logs = []
    assert session.logout(page, lambda m, *a: logs.append(m)) is True
    assert logs == ["Logged out of the URA portal."]


def test_a_home_page_that_will_not_open_still_ends_quietly(page, monkeypatch):
    page.set_content("<html><body><p>some other screen</p></body></html>")
    monkeypatch.setattr(eng, "ensure_portal_home", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("no menu")))
    logs = []
    assert session.logout(page, lambda m, *a: logs.append(m)) is False
    assert logs[-1] == "Could not find the Logout button - the browser is being closed anyway."
