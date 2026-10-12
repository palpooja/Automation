"""User, 6 Oct 2026: in 'Open portal (manual)' (used for screenshots and page captures) nothing could
be clicked after login and the portal's second frame never appeared (10.6 and 10.7 alike).

Cause, reproduced: with 'Load portal pages faster' ticked, every request the browser makes is passed
to the program's filter (ledger_engine.speed_up, a Playwright route), which only answers while the
program is working with the browser.  During login the program checks the browser every 0.5 s, so
login works; after the hand-over it only waited for a button press, so every request was held.  The
same applied to hand_over() during a run.  Fix: manual mode switches the filter off for the person's
browser, and every wait for the person keeps the browser serviced (BROWSER_KEEPALIVE).

The browser runs on a worker thread here, as in app.py.
"""
import contextlib
import queue
import threading

import pytest

from core import manual, session
from core.runner import RunContext
from engines import ledger_engine as eng
from tests.auto.conftest import launch_chromium

pytestmark = pytest.mark.browser
CLIENT = {"name": "Test client", "login": "1000000001", "password": "not-a-real-password"}


@pytest.fixture
def setup(monkeypatch, mock_url, tmp_path):
    monkeypatch.setattr(eng, "SPEED_UPS", True)                       # as on the user's PC
    monkeypatch.setattr(eng, "launch_browser", lambda p: launch_chromium(p))
    monkeypatch.setattr(eng, "login", lambda page, tin, pw: page.goto(f"{mock_url}/Portal/after_login.html"))
    monkeypatch.setattr(session, "logout", lambda page, log: None)
    seen = {}

    @contextlib.contextmanager
    def watched(ctx, client):
        with session.open_session(ctx, client) as page:
            yield page
            # the page's own clock, before the browser closes: did the frame load during the wait?
            seen["t"] = page.evaluate("() => [window.t0, window.loadedAt || null, Date.now()]")
            seen["keepalive_inside"] = getattr(eng, "BROWSER_KEEPALIVE", None) is not None
    reply_q = queue.Queue()
    ctx = RunContext(queue.Queue(), reply_q, threading.Event(), tmp_path, tmp_path / "log.txt")
    return ctx, reply_q, watched, seen


def _in_worker(fn):
    err = []

    def run():
        try:
            fn()
        except BaseException as e:            # pragma: no cover - reported below
            err.append(e)
    t = threading.Thread(target=run)
    t.start()
    t.join(60)
    if err:
        raise err[0]


def _loaded_during_wait(seen):
    t0, loaded, now = seen["t"]
    return loaded is not None and (now - loaded) / 1000 > 1.0


def test_manual_mode_browser_keeps_loading_after_login(setup):
    ctx, reply_q, watched, seen = setup
    threading.Timer(4, lambda: reply_q.put("")).start()     # the person presses 'Close browser'
    _in_worker(lambda: manual.run_manual(ctx, CLIENT, watched))
    assert _loaded_during_wait(seen), "the second frame must load while the browser is the person's"


def test_hand_over_during_a_run_keeps_the_browser_loading(setup):
    ctx, reply_q, watched, seen = setup
    threading.Timer(4, lambda: reply_q.put("c")).start()

    def run():
        with watched(ctx, CLIENT):
            eng.hand_over("set something in the browser")
    _in_worker(run)
    assert _loaded_during_wait(seen), "the browser must keep loading while it waits for the person"
    assert seen["keepalive_inside"]
    assert getattr(eng, "BROWSER_KEEPALIVE", None) is None, "cleared when the session ends"
