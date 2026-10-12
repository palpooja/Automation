"""Returns, live run of 5 Oct 2026 for client B (18:15): 'Turning to page 2...' then 13 s later
"'Next' was clicked but the table did not change - treating this as the last page", and the
warning "portal reported 12 record(s) but only 6 file(s) are on disk".  Page turns in the live
logs took 6-20 s; the program allowed about 12 s, whatever the 'wait until it downloads' setting."""
import pytest

pytestmark = pytest.mark.browser


def test_a_slow_page_turn_is_waited_for(page, mock_url, tmp_path, monkeypatch):
    from engines import returns_common, returns_pw
    monkeypatch.setattr(returns_common, "TAB_WAIT_SECONDS", 5)
    monkeypatch.setattr(returns_common, "PAGE_WAIT_SECONDS", 5)
    logs = []
    page.goto(f"{mock_url}/home.html")
    page.evaluate("() => sessionStorage.setItem('nextDelayMs', '15000')")     # page 2 needs 15 s
    res = {}
    returns_pw.run_on_page(page, "1000012345", "01/07/2025", "03/10/2026", ["Monthly VAT Return"],
                           lambda m, *a: logs.append(m), lambda: None, str(tmp_path), patient=False,
                           result=res, skip_existing=True)
    assert not any("treating this as the last page" in m for m in logs), logs
    assert any("Page 2: 1 return(s) on this page." in m for m in logs), logs
    # the page-2 return (R-VAT-BAD never loads in the mock) was at least tried, so it is reported
    assert any(i["status"] == "failed" and "page 2" in i["item"] for i in res["items"]), res["items"]
