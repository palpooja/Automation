"""Chrome window closed by hand during a run (live log 6 Oct 2026, 20:36): Returns ended with
TargetClosedError and the words 'Stopped - writing the status sheet'; Ledgers and Certificates were
still started on the dead browser and each 'failed'; no clear message.  Wanted: stop the run cleanly."""
import time

import pytest
from playwright._impl._errors import TargetClosedError

from core import session
from engines import ledger_engine as eng, returns_pw

pytestmark = pytest.mark.browser


def test_browser_gone_is_recognised(page):
    assert session.browser_gone(page) is False
    assert session.browser_gone(None, TargetClosedError()) is True
    assert session.browser_gone(None, RuntimeError("Level 3 box not found")) is False
    page.close()
    assert session.browser_gone(page) is True


def test_ledger_level1_wait_ends_at_once_on_a_closed_browser(page, monkeypatch):
    monkeypatch.setattr(eng, "SLOW_FACTOR", 0.1)           # 90 s x 0.1 = 9 s before the fix
    page.close()
    t0 = time.monotonic()
    assert eng.wait_for_level1_patiently(page) is False
    assert time.monotonic() - t0 < 1.0


def test_ledger_step_ends_at_once_when_the_browser_is_closed(tmp_path, monkeypatch):
    calls = []

    def process(page, tax, *a, **k):
        calls.append(tax)
        raise TargetClosedError()
    for name in ("open_tax_statement", "debug_dump", "dismiss_popup_if_present", "safe_screenshot",
                 "recover_to_level1", "write_summary", "update_tax_head_files"):
        monkeypatch.setattr(eng, name, lambda *a, **k: None)
    monkeypatch.setattr(eng, "process_tax_type", process)
    monkeypatch.setattr(eng, "OUTPUT_ROOT", str(tmp_path))
    monkeypatch.setattr(eng, "ONLINE_WAIT", None)
    with pytest.raises(session.BrowserClosed):
        eng.run_job_on_page(object(), {"tin": "1000000001", "password": "x", "period_from": "2025-07",
                                       "period_to": "2026-06", "tax_types": ["Value Added Tax", "PAYE"],
                                       "other_jobs": []})
    assert calls == ["Value Added Tax"], "no 3 attempts and no further tax types on a dead browser"


def test_returns_says_the_browser_was_closed(page, mock_url, tmp_path, monkeypatch):
    from engines import returns_common
    monkeypatch.setattr(returns_common, "TAB_WAIT_SECONDS", 5)
    real = returns_pw.first_row_text

    def close_then(frame):
        page.close()                                        # the person closes Chrome
        return real(frame)
    monkeypatch.setattr(returns_pw, "first_row_text", close_then)
    logs = []
    page.goto(f"{mock_url}/home.html")
    with pytest.raises(Exception) as err:
        returns_pw.run_on_page(page, "1000012345", "01/07/2025", "03/10/2026", ["Monthly VAT Return"],
                               lambda m, *a: logs.append(m), lambda: None, str(tmp_path), patient=False,
                               result={}, skip_existing=True)
    assert session.browser_gone(None, err.value)
    assert any("The browser window was closed" in m for m in logs), logs
    assert not any(m.startswith("Stopped - ") for m in logs), "not reported as a Stop"
