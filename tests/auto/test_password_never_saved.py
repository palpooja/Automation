"""Finding H6: a failed login must never put the password into the log, the history or the summaries."""
import queue
import threading

import pytest

from core import store
from core.runner import RunContext
from engines import ledger_engine as eng

PASSWORD = "Real-Client-Password-42"
HIDDEN_BOX = ('<html><body><input id="userName" type="text">'
              '<input type="password" style="display:none"><button>Log In</button></body></html>')


@pytest.mark.browser
def test_fill_error_message_does_not_contain_the_password(page, monkeypatch):
    page.set_content(HIDDEN_BOX)
    page.set_default_timeout(1500)
    with pytest.raises(Exception) as info:
        eng.fill_login_form(page, "1000000002", PASSWORD)
    assert PASSWORD not in str(info.value)
    assert PASSWORD not in repr(info.value)


def test_run_context_masks_known_passwords_in_log_file_and_items(tmp_path):
    ctx = RunContext(queue.Queue(), queue.Queue(), threading.Event(), tmp_path, tmp_path / "run.txt")
    ctx.add_secret(PASSWORD)
    ctx.log(f'Call log: fill("{PASSWORD}")')
    ctx.record("login / browser", "failed", f'fill("{PASSWORD}")')
    ctx.close()
    assert PASSWORD not in (tmp_path / "run.txt").read_text(encoding="utf-8")
    assert PASSWORD not in ctx.items[-1]["detail"]
    assert PASSWORD not in ctx.scrub(f"x {PASSWORD} y")


def test_history_never_stores_a_known_password(fake_data_dir):
    store.add_history({"note": f'fill("{PASSWORD}")'}, secrets=[PASSWORD])
    assert PASSWORD not in store.HISTORY_FILE.read_text(encoding="utf-8")
