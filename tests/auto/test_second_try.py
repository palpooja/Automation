"""10.15 - user, 8 Oct 2026: 'Put Retry if anything else is failed like returns.'

Like the slow returns of 10.14, anything that failed during a client's steps is tried once more after the
client's other steps: a whole step that stopped with an error, a ledger tax type / month / year, a WHT
certificate type with a failed month, a PAYE year, a return type with a return not downloaded.  The second
try UPDATES the first try's items on the Review page ('saved on second try'); it never lists them twice and
never puts anything off a third time.  No portal here - the steps are stand-ins; invented logins only."""
import queue
import threading

from core.runner import RunContext, needs_review


def _ctx(tmp_path):
    ctx = RunContext(queue.Queue(), queue.Queue(), threading.Event(), tmp_path, tmp_path / "run.log")
    ctx.set_scope("Client A", "1000000001", "Ledgers")
    return ctx


def test_second_try_updates_the_first_try_items(tmp_path):
    ctx = _ctx(tmp_path)
    ctx.record("Value Added Tax ledger: MAY-2026", "failed", "Level 3 would not open")
    ctx.record("Value Added Tax ledger: JUN-2026", "saved")
    ctx.record("Value Added Tax ledger: whole tax type", "failed", "Timeout")
    ctx.final_try = True
    ctx.set_scope("Client A", "1000000001", "Ledgers - second try for Value Added Tax")
    ctx.record("Value Added Tax ledger: MAY-2026", "saved", "Payment 1,000")
    ctx.record("Value Added Tax ledger: JUN-2026", "already downloaded")
    ctx.record("Value Added Tax ledger: JUL-2026", "saved")                     # new: listed
    assert ctx.resolve("Value Added Tax ledger: whole tax type", "done on second try")
    got = {it["item"]: it["status"] for it in ctx.items}
    assert len(ctx.items) == 4, "nothing listed twice"
    assert got["Value Added Tax ledger: MAY-2026"] == "saved on second try"
    assert got["Value Added Tax ledger: JUN-2026"] == "saved", "already on disk: the first try's line stays"
    assert got["Value Added Tax ledger: whole tax type"] == "done on second try"
    assert not any(needs_review(s) for s in got.values())


def test_still_failing_keeps_the_failure_with_the_second_reason(tmp_path):
    ctx = _ctx(tmp_path)
    ctx.record("PAYE certificates July 2025 to June 2026", "failed", "no results within 60 s")
    ctx.final_try = True
    ctx.record("PAYE certificates July 2025 to June 2026", "failed", "no results again")
    (it,) = ctx.items
    assert it["status"] == "failed" and it["detail"] == "second try: no results again"


def test_ledgers_retry_only_what_failed():
    from modules.ledger import LedgerModule
    results = [{"Section": "Value Added Tax ledger", "Item": "MAY-2026", "Status": "failed"},
               {"Section": "PAYE ledger", "Item": "MAY-2026", "Status": "saved"},
               {"Section": "PAYE ledger", "Item": "JUN-2026", "Status": "check"},        # a difference, not a failure
               {"Section": "Income Tax ledger", "Item": "whole tax type", "Status": "failed"},
               {"Section": "taxpayer profile", "Item": "taxpayer profile", "Status": "failed"}]
    settings = {"tax_types": ["Value Added Tax", "PAYE", "Income Tax"], "other_jobs": ["profile"]}
    assert LedgerModule.failed_work(results, settings) == (["Value Added Tax", "Income Tax"], ["profile"])


def test_returns_retry_types_with_failures_but_not_ones_skipped_by_the_person():
    from modules.returns import ReturnsModule
    items = [{"item": "Monthly VAT Return: 3. September 2025", "status": "failed",
              "detail": "not downloaded after all retries"},
             {"item": "PAYE Return", "status": "failed", "detail": "no results area appeared"},
             {"item": "Withholding Tax Return: 1. July 2025", "status": "failed",
              "detail": "skipped by you while waiting - tried again on the next run"},
             {"item": "Monthly VAT Return: 4. October 2025", "status": "saved"}]
    assert ReturnsModule.failed_types(items) == ["Monthly VAT Return", "PAYE Return"]


def test_certificates_second_try_fetches_only_the_failed_paye_year(tmp_path, monkeypatch):
    from core import session
    from engines import paye_engine
    from modules.wht_receivable import WhtReceivableModule
    calls = []

    def fake_run(page, tin, pf, pt, out, log, check, record, failed=None, years=None, **kw):
        calls.append(years)
        if years is None:                      # first try: one year gets no answer
            record("PAYE certificates July 2025 to June 2026", "failed", "no results within 60 s")
            failed.append("2025-2026")
        else:
            record("PAYE certificates July 2025 to June 2026 - workbook", "saved", "2 employee(s)")
        return []
    monkeypatch.setattr(paye_engine, "run", fake_run)
    monkeypatch.setattr(session, "go_home", lambda *a, **k: True)
    ctx = _ctx(tmp_path)
    opts = {"generate_for": [], "paye": True, "period": ("2025-07", "2027-06")}
    mod = WhtReceivableModule.__new__(WhtReceivableModule)
    mod.run(ctx, {"login": "1000000001", "password": "x"}, opts, page=object())
    (title, again), = ctx.after_client
    assert title == "Certificates - second try for PAYE July 2025 to June 2026"
    ctx.after_client, ctx.final_try = [], True
    again(object())
    assert calls == [None, ["2025-2026"]], "the second try asks for the failed year only"
    assert ctx.after_client == [], "never put off a third time"
