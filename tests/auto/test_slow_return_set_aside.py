"""10.14 (user, 8 Oct 2026): a VAT return never opened; the run waited 8.5 min until Chrome was closed, and the
next return, Ledgers and Certificates were never done.  Now:
1. first try: a return not ready within FIRST_TRY_SECONDS is set aside, the run goes on;
2. second try, after the client's other steps, with the mode's full wait;
3. during that wait, after SKIP_OFFER_AFTER_SECONDS: the voice says 'Attention needed' and the yellow bar offers
   Skip / Keep waiting - the wait goes on while it shows."""
import time

import pytest

from engines import returns_common as rc

pytestmark = pytest.mark.browser


class Person:
    """Stands in for the window: answers after `after` s with `answer`."""
    def __init__(self, answer=None, after=0.5):
        self.answer, self.after, self.asked, self.at, self.hidden = answer, after, [], None, 0

    def attention(self, message):
        self.asked.append(message)
        self.at = time.monotonic()

    def poll_answer(self):
        if self.answer and self.at and time.monotonic() - self.at >= self.after:
            a, self.answer = self.answer, None
            return a
        return None

    def hide_question(self):
        self.hidden += 1


def _run(page, mock_url, tmp_path, delay_ms, **kw):
    from engines import returns_pw
    page.goto(f"{mock_url}/home.html")
    page.evaluate(f"() => sessionStorage.setItem('openDelayMs', '{delay_ms}')")
    res, logs = {}, []
    t0 = time.monotonic()
    returns_pw.run_on_page(page, "1000012345", "01/07/2025", "30/06/2026", ["Rental Income Tax Return"],
                           logs.append, lambda: None, str(tmp_path), patient=False, result=res,
                           skip_existing=True, **kw)
    return res, logs, time.monotonic() - t0


def test_first_try_sets_a_slow_return_aside_and_goes_on(page, mock_url, tmp_path, monkeypatch):
    monkeypatch.setattr(rc, "TAB_WAIT_SECONDS", 60)
    monkeypatch.setattr(rc, "PAGE_WAIT_SECONDS", 60)
    res, logs, took = _run(page, mock_url, tmp_path, 30000, first_try_s=3)
    assert [d["label"] for d in res["deferred"]] == ["2. August 2025"]
    assert any(i["status"] == "set aside" for i in res["items"])
    folder = tmp_path / "1000012345" / "Rental Income Tax Return"
    assert (folder / "1. July 2025.pdf").exists(), "the next return is not held up"
    assert took < 25, f"{took:.0f} s - must not wait for the slow return"
    assert not any(i["status"] in ("failed", "incomplete") for i in res["items"]), res["items"]


def test_second_try_waits_fully_and_saves_it(page, mock_url, tmp_path, monkeypatch):
    monkeypatch.setattr(rc, "TAB_WAIT_SECONDS", 30)
    monkeypatch.setattr(rc, "PAGE_WAIT_SECONDS", 30)
    _run(page, mock_url, tmp_path, 6000, first_try_s=2)                     # set aside
    res, logs, took = _run(page, mock_url, tmp_path, 6000, first_try_s=None, person=Person())
    assert (tmp_path / "1000012345" / "Rental Income Tax Return" / "2. August 2025.pdf").exists()
    assert res["deferred"] == []


def test_second_try_attention_and_skip(page, mock_url, tmp_path, monkeypatch):
    monkeypatch.setattr(rc, "TAB_WAIT_SECONDS", 120)
    monkeypatch.setattr(rc, "PAGE_WAIT_SECONDS", 120)
    monkeypatch.setattr(rc, "SKIP_OFFER_AFTER_SECONDS", 1)
    person = Person("s", after=0.5)
    res, logs, took = _run(page, mock_url, tmp_path, 60000, first_try_s=None, person=person)
    assert len(person.asked) == 1 and "2. August 2025" in person.asked[0] and "Skip" in person.asked[0]
    assert took < 30, f"{took:.0f} s - Skip ends the wait"
    assert any(i["status"] == "failed" and "skipped by you" in i["detail"] for i in res["items"])
    assert (tmp_path / "1000012345" / "Rental Income Tax Return" / "1. July 2025.pdf").exists()


def test_second_try_keep_waiting_goes_on_waiting(page, mock_url, tmp_path, monkeypatch):
    monkeypatch.setattr(rc, "TAB_WAIT_SECONDS", 60)
    monkeypatch.setattr(rc, "PAGE_WAIT_SECONDS", 60)
    monkeypatch.setattr(rc, "SKIP_OFFER_AFTER_SECONDS", 1)
    person = Person("w", after=0.5)
    res, logs, took = _run(page, mock_url, tmp_path, 5000, first_try_s=None, person=person)
    assert len(person.asked) == 1 and person.hidden == 1
    assert (tmp_path / "1000012345" / "Rental Income Tax Return" / "2. August 2025.pdf").exists()


def test_attention_says_it_and_shows_skip_without_waiting(tmp_path, monkeypatch):
    import queue
    import threading
    from core import voice
    from core.runner import RunContext
    said = []
    monkeypatch.setattr(voice, "say", lambda parts, choice="1": said.append(parts) or True)
    ui, rq = queue.Queue(), queue.Queue()
    ctx = RunContext(ui, rq, threading.Event(), tmp_path, tmp_path / "log.txt")
    ctx.voice_choice = "1"
    t0 = time.monotonic()
    ctx.attention("VAT: 2. August 2025 is still not open")
    assert time.monotonic() - t0 < 0.5, "never waits for the person"
    assert said == [["Attention needed."]]
    msgs = []
    while not ui.empty():
        msgs.append(ui.get())
    assert any(k == "ask" and p[1] == "skipwait" for k, p in msgs)
    assert ctx.poll_answer() is None
    rq.put("s")
    assert ctx.poll_answer() == "s"
