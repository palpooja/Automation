"""Shared test fixtures.

* PROJECT root is put on sys.path, so `import core`, `import engines` work from any folder.
* `mock_url`  - serves tests/mock_portal on a free local port for the whole session.
* `browser`   - one headless Chromium for the whole session. It uses, in order:
      1. the executable in the URA_TEST_CHROMIUM environment variable, if set;
      2. Playwright's own Chromium (run once:  python -m playwright install chromium);
      3. the installed Google Chrome.
  Tests marked `browser` are skipped (with the reason shown) if none of these starts.
"""
import functools
import http.server
import os
import socketserver
import sys
import threading
from pathlib import Path

import pytest

PROJECT = Path(__file__).resolve().parents[2]
MOCK_DIR = PROJECT / "tests" / "mock_portal"
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))

LAUNCH_ARGS = ["--no-sandbox", "--disable-gpu", "--disable-dev-shm-usage"]


class _QuietHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *args):          # keep test output readable
        pass

    def do_POST(self):
        import sys as _s
        _s.path.insert(0, str(MOCK_DIR))
        import ledger_server
        import paye_server
        if not (ledger_server.handle(self, "POST") or paye_server.handle(self, "POST")):
            self.send_error(404)

    def do_GET(self):
        import sys as _s
        _s.path.insert(0, str(MOCK_DIR))
        # 10.18: the WHT certificate page's live address, answered after ?ms= (a slow link); the page itself
        # is the test copy Portal/wht_live.html.  Before the PAYE handler, which owns the same controller path.
        if "loadWthPage" in self.path:
            import time
            from urllib.parse import parse_qs, urlparse
            ms = int(parse_qs(urlparse(self.path).query).get("ms", ["0"])[0])
            time.sleep(ms / 1000)
            body = (MOCK_DIR / "Portal" / "wht_live.html").read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        import ledger_server
        if ledger_server.handle(self, "GET"):          # the Ledger form screens (option C tests)
            return
        import paye_server
        if paye_server.handle(self, "GET"):            # the PAYE certificate screens (live capture 8 Oct 2026)
            return
        # /slow?ms=7000 answers after a delay: a portal still fetching its schedules
        if self.path.startswith("/slow"):
            import time
            from urllib.parse import parse_qs, urlparse
            ms = int(parse_qs(urlparse(self.path).query).get("ms", ["1000"])[0])
            time.sleep(ms / 1000)
            body = b"ok"
            self.send_response(200)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        super().do_GET()


@pytest.fixture(scope="session")
def mock_url():
    handler = functools.partial(_QuietHandler, directory=str(MOCK_DIR))
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.ThreadingTCPServer(("127.0.0.1", 0), handler) as srv:
        port = srv.server_address[1]
        t = threading.Thread(target=srv.serve_forever, daemon=True)
        t.start()
        yield f"http://127.0.0.1:{port}"
        srv.shutdown()


def launch_chromium(p):
    exe = os.environ.get("URA_TEST_CHROMIUM")
    attempts = ([{"executable_path": exe}] if exe else []) + [{}, {"channel": "chrome"}]
    last = None
    for kw in attempts:
        try:
            return p.chromium.launch(headless=True, args=LAUNCH_ARGS, **kw)
        except Exception as e:          # try the next way to start a browser
            last = e
    pytest.skip(f"No Chromium available for browser tests ({type(last).__name__}). "
                "Set URA_TEST_CHROMIUM or run: python -m playwright install chromium")


@pytest.fixture(scope="session")
def playwright_instance():
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        yield p


@pytest.fixture(scope="session")
def browser(playwright_instance):
    b = launch_chromium(playwright_instance)
    yield b
    b.close()


@pytest.fixture
def page(browser):
    ctx = browser.new_context(accept_downloads=True)
    pg = ctx.new_page()
    yield pg
    ctx.close()


@pytest.fixture
def fake_data_dir(tmp_path, monkeypatch):
    """Points core.store at an empty temporary folder (never the real %LOCALAPPDATA%)."""
    from core import store
    monkeypatch.setattr(store, "DATA_DIR", tmp_path)
    monkeypatch.setattr(store, "LOG_DIR", tmp_path / "logs")
    monkeypatch.setattr(store, "CLIENTS_FILE", tmp_path / "clients.json")
    monkeypatch.setattr(store, "HISTORY_FILE", tmp_path / "history.json")
    monkeypatch.setattr(store, "SETTINGS_FILE", tmp_path / "settings.json")
    monkeypatch.setattr(store, "OLD_RETURNS_CLIENTS", tmp_path / "no_old_file.json")
    return tmp_path


@pytest.fixture
def ledger_portal():
    """The test copy of the Ledger form screens (tests/mock_portal/ledger_server.py), reset."""
    import sys as _s
    _s.path.insert(0, str(MOCK_DIR))
    import ledger_server
    ledger_server.reset()
    from engines import ledger_engine as eng
    eng.DIRECT_STATE.clear()            # 10.17: the route proof is per client - never carried between tests
    eng.ROUTE_STATE.clear()
    eng.RESULTS.clear()                 # 10.18: nor the items of an earlier test's tax type
    eng.NOTES.clear()
    return ledger_server


@pytest.fixture
def paye_portal():
    """The test copy of the live PAYE certificate screens (tests/mock_portal/paye_server.py), reset."""
    import sys as _s
    _s.path.insert(0, str(MOCK_DIR))
    import paye_server
    paye_server.reset()
    return paye_server
