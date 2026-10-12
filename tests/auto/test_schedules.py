"""Returns and 'Print with Schedules' (user, 4 Oct 2026: VAT schedules missing; then: URA no longer
gives VAT schedules).  The PDF is made when the portal asks to print; if it does not ask, when the page
has been quiet (no network, no change) for as long as the old fixed wait; never later than a cap."""
import time

import pytest
from pypdf import PdfReader

from engines import returns_pw

pytestmark = pytest.mark.browser


@pytest.fixture
def ctx(browser):
    c = browser.new_context()
    returns_pw.install_print_catcher(c)
    yield c
    c.close()


def _run(ctx, mock_url, mode, limit=20, tmp_path=None, expects_schedules=True):
    tab = ctx.new_page()
    tab.goto(f"{mock_url}/Portal/schedules_test.html?mode={mode}")
    t0 = time.monotonic()
    page_, detail = returns_pw.press_print_with_schedules(ctx, tab, tab.locator("#printWthSchid"), limit,
                                                         expects_schedules=expects_schedules)
    took = time.monotonic() - t0
    text = ""
    if tmp_path is not None:
        out = str(tmp_path / f"{mode}.pdf")
        returns_pw._print_to_pdf(page_, out)
        text = " ".join(p.extract_text() or "" for p in PdfReader(out).pages)
    return detail, took, text


def test_portal_ready_at_once_like_paye_no_fixed_wait(ctx, mock_url, tmp_path):
    detail, took, text = _run(ctx, mock_url, "fast", tmp_path=tmp_path)
    assert detail == returns_pw.READY and "SCHEDULES FOR THIS RETURN" in text
    assert took < 2, f"{took:.1f} s (the old fixed wait was about 5.5 s)"


def test_schedules_still_loading_from_the_portal_are_waited_for(ctx, mock_url, tmp_path):
    detail, took, text = _run(ctx, mock_url, "slow", tmp_path=tmp_path)
    assert detail == returns_pw.READY and "SCHEDULES FOR THIS RETURN" in text and took >= 7


def test_schedules_in_a_new_tab_are_printed_from_that_tab(ctx, mock_url, tmp_path):
    detail, _, text = _run(ctx, mock_url, "newtab", tmp_path=tmp_path)
    assert detail.startswith(returns_pw.READY) and "SCHEDULES FOR THIS RETURN" in text


def test_vat_with_no_schedules_from_ura_is_printed_after_the_old_wait_and_not_flagged(ctx, mock_url):
    detail, took, _ = _run(ctx, mock_url, "never", expects_schedules=False)
    assert not detail.startswith("check") and "URA" in detail
    assert returns_pw.QUIET_S <= took < returns_pw.QUIET_S + 2


def test_other_return_types_without_schedules_are_flagged_for_a_person(ctx, mock_url):
    detail, took, _ = _run(ctx, mock_url, "never", expects_schedules=True)
    assert detail.startswith("check:") and took < returns_pw.QUIET_S + 2


def test_a_page_that_never_settles_stops_at_the_limit_and_is_flagged(ctx, mock_url):
    detail, took, _ = _run(ctx, mock_url, "busy", limit=8)
    assert detail.startswith("check:") and 8 <= took < 10


def test_monthly_vat_is_the_type_ura_no_longer_gives_schedules_for():
    assert "Monthly VAT Return" in returns_pw.NO_SCHEDULES_FROM_URA
    assert "PAYE Return" not in returns_pw.NO_SCHEDULES_FROM_URA
