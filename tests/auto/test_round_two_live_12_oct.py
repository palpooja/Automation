"""10.18 - three things the live run of 12 Oct 2026 (Tronics, slow link, build 10.16) showed:
* Returns: the VAT search answered nothing for 7.5 min, and every later return type was skipped with 'no date
  boxes' - the page was never opened again;
* Certificates: the WHT frame was at the certificate page's address but its form took longer than 12 s; every
  re-click restarted the load and after 40 s the whole step failed in one second;
* Ledgers: one month's export failed and the WHOLE tax type started again from Level 1 (4.5 min).
Test pages: tests/mock_portal/Portal/history.html (flags), Portal/wht_menu_slow.html, ledger_server.py."""
import os
import time

import pytest

from engines import ledger_engine as eng, returns_common, returns_pw, wht_engine as we

pytestmark = pytest.mark.browser


def _returns(page, mock_url, tmp_path, monkeypatch, wanted, flags):
    monkeypatch.setattr(returns_common, "TAB_WAIT_SECONDS", 5)
    monkeypatch.setattr(returns_common, "PAGE_WAIT_SECONDS", 5)
    page.goto(f"{mock_url}/home.html")
    page.evaluate("f => { sessionStorage.clear(); for (const k in f) sessionStorage.setItem(k, f[k]); }", flags)
    logs, res, reopened = [], {}, []

    def reopen():
        reopened.append(1)
        page.goto(f"{mock_url}/home.html")
    returns_pw.run_on_page(page, "1000012345", "01/07/2025", "03/10/2026", wanted, logs.append, lambda: None,
                           str(tmp_path), patient=False, result=res, skip_existing=True, reopen=reopen)
    return logs, res, reopened


def test_a_search_that_answers_nothing_does_not_cost_the_later_return_types(page, mock_url, tmp_path, monkeypatch):
    logs, res, reopened = _returns(page, mock_url, tmp_path, monkeypatch,
                                   ["Monthly VAT Return", "Local Excise Duty Return"], {"deadSearch": "VATRT"})
    by = {i["item"]: i for i in res["items"] if i["status"] in ("failed", "saved")}
    assert by["Monthly VAT Return"]["detail"] == "no results area appeared"
    saved = [i for i in res["items"] if i["status"] == "saved" and i["item"].startswith("Local Excise")]
    assert len(saved) == 1, logs
    assert reopened == [1], "the home page once, then 'View return history' again"
    assert sum("Opening 'View return history'" in m for m in logs) == 2, logs
    copies = [f for f in os.listdir(tmp_path / "1000012345") if f.startswith("Search failed Monthly VAT Return")]
    assert any(f.endswith(".png") for f in copies) and any(f.endswith(".html") for f in copies), copies


def test_missing_date_boxes_are_recovered_once_from_the_home_page(page, mock_url, tmp_path, monkeypatch):
    # the portal's order is VAT, PAYE, Excise: PAYE's page loses its date boxes after its search, Excise follows
    logs, res, reopened = _returns(page, mock_url, tmp_path, monkeypatch,
                                   ["PAYE Return", "Local Excise Duty Return"], {"dropBoxesAfter": "PAYERT"})
    assert reopened == [1], logs
    saved = [i for i in res["items"] if i["status"] == "saved" and i["item"].startswith("Local Excise")]
    assert len(saved) == 1, logs                                          # searched after the recovery
    assert not any("Skipping this category" in m for m in logs), logs
    assert any("No date boxes on the page - opening the return history page again" in m for m in logs), logs


def test_a_certificate_page_on_its_way_is_waited_for_not_clicked_again(page, mock_url, monkeypatch):
    monkeypatch.setattr(we, "PAGE_OPEN_S", 40)
    page.goto(f"{mock_url}/Portal/wht_menu_slow.html")
    page.evaluate("() => sessionStorage.setItem('whtMs', '16000')")     # the form shows 16 s after the click
    logs = []
    bot = we.WhtBot(page, lambda m, *a: logs.append(m), lambda: None)
    t0 = time.monotonic()
    bot.open_wht_page()
    took = time.monotonic() - t0
    assert page.evaluate("() => window.clicks") == 1, "the menu link was pressed once"
    assert any("certificate page is loading" in m for m in logs), logs
    assert logs[-1] == "Search form is open."
    assert 15 < took < 35, f"{took:.0f} s"


def test_a_failed_month_is_recorded_for_that_month_and_the_next_month_goes_on(page, mock_url, ledger_portal, tmp_path,
                                                                           monkeypatch, capsys):
    import pandas as pd
    page.goto(f"{mock_url}{ledger_portal.BASE}loadLevelTwo")
    monkeypatch.setattr(eng, "open_level2", lambda pg, *a, **k: pg.main_frame)
    monkeypatch.setattr(eng, "recover_to_level1", lambda *a, **k: None)
    monkeypatch.setattr(eng, "_back_to_level1_after", lambda *a, **k: None)
    monkeypatch.setattr(eng, "export_level2_to_excel", lambda *a, **k: "portal export")
    monkeypatch.setattr(eng, "SKIP_NIL", False)
    monkeypatch.setattr(eng, "read_level2_table", lambda fr: pd.DataFrame(
        {"Period": [m for m, _ in ledger_portal.MONTHS], "Payment": [a for _, a in ledger_portal.MONTHS]}))
    reloaded = []
    monkeypatch.setattr(eng, "refresh_level2", lambda pg, fr, tax: (reloaded.append(tax), pg.goto(
        f"{mock_url}{ledger_portal.BASE}loadLevelTwo"), pg.main_frame)[2])
    real = eng.level3_month

    def flaky(pg, fr, folder, tax, month, payment_value=None):
        if month == "MAY-2026":
            raise RuntimeError("Locator.evaluate: Timeout 10000ms exceeded.")
        return real(pg, fr, folder, tax, month, payment_value)
    monkeypatch.setattr(eng, "level3_month", flaky)
    eng.RESULTS.clear()
    eng.process_tax_type(page, "PAYE", "2026-04", "2026-06", tmp_path, "1000000001", "pw")
    status = {r["Item"]: r["Status"] for r in eng.RESULTS if r["Section"] == "PAYE ledger"}
    assert status["MAY-2026"] == "failed" and status["JUN-2026"] == "saved" and status["APR-2026"] == "saved", status
    assert reloaded == ["PAYE"], "Level 2 reloaded once, after the failed month"
    out = capsys.readouterr().out
    assert "MAY-2026: failed (Locator.evaluate: Timeout 10000ms exceeded.) - recorded for this month" in out, out
