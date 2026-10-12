"""Findings H1 (stored passwords lost), H2 (another client's password used), L1 (Login ID checks)."""
import base64
import json

import pytest

from core import store
from engines import ledger_engine as eng

UNDECRYPTABLE = base64.b64encode(b"\xff\xfe-not-decryptable").decode()


def _write(path, data):
    path.write_text(json.dumps(data), encoding="utf-8")


def test_unreadable_password_is_flagged_and_its_original_text_survives_the_next_save(fake_data_dir):
    _write(store.CLIENTS_FILE, [
        {"name": "A", "login": "1000000001", "password": store.protect_secret("pw-A")},
        {"name": "B", "login": "1000000002", "password": UNDECRYPTABLE}])
    clients = store.load_clients()
    assert clients[0]["password"] == "pw-A"
    assert clients[1]["password"] == ""
    assert clients[1].get("password_unreadable") is True
    clients.append({"name": "C", "login": "1000000003", "password": "pw-C"})   # Clients > Add
    store.save_clients(clients)
    after = json.loads(store.CLIENTS_FILE.read_text(encoding="utf-8"))
    assert after[1]["password"] == UNDECRYPTABLE, "the only copy of B's password was overwritten"
    assert store.unprotect_secret(after[2]["password"]) == "pw-C"


def test_re_entering_an_unreadable_password_replaces_it(fake_data_dir):
    _write(store.CLIENTS_FILE, [{"name": "B", "login": "1000000002", "password": UNDECRYPTABLE}])
    clients = store.load_clients()
    clients[0] = {"name": "B", "login": "1000000002", "password": "new-pw"}       # Clients > Edit
    store.save_clients(clients)
    assert store.load_clients()[0]["password"] == "new-pw"


def test_damaged_clients_file_is_set_aside_not_overwritten(fake_data_dir):
    damaged = '[{"name": "A", "login": "1000000001", "password": "x"}, {"name": "B"'
    store.CLIENTS_FILE.write_text(damaged, encoding="utf-8")
    clients = store.load_clients()
    assert clients == []
    kept = list(fake_data_dir.glob("clients.damaged-*.json"))
    assert len(kept) == 1 and kept[0].read_text(encoding="utf-8") == damaged
    assert store.LOAD_WARNINGS, "the window must be told the file was damaged"
    store.save_clients([{"name": "D", "login": "1000000004", "password": "pw-D"}])
    assert kept[0].read_text(encoding="utf-8") == damaged


def test_previous_client_list_is_kept_as_a_backup_on_every_save(fake_data_dir):
    store.save_clients([{"name": "A", "login": "1000000001", "password": "pw-A"}])
    store.save_clients([{"name": "A", "login": "1000000001", "password": "pw-A"},
                        {"name": "B", "login": "1000000002", "password": "pw-B"}])
    backup = json.loads(store.CLIENTS_FILE.with_name("clients.json.bak").read_text(encoding="utf-8"))
    assert [c["name"] for c in backup] == ["A"]


@pytest.fixture
def capture_login(monkeypatch):
    sent = {}
    monkeypatch.setattr(eng, "open_portal", lambda page: True)
    monkeypatch.setattr(eng, "fill_login_form",
                        lambda page, tin, password: sent.update(tin=tin, password=password))
    monkeypatch.setattr(eng, "ask_user", lambda *a, **k: "q")
    monkeypatch.setattr(eng, "CREDS", {"tin": None, "password": None})
    return sent


def test_login_never_sends_another_clients_password(capture_login):
    eng.CREDS.update(tin="1000000001", password="pw-A")          # left over from client A
    with pytest.raises(RuntimeError):
        eng.login(object(), "1000000002", "")                      # client B, password unreadable
    assert capture_login.get("password") != "pw-A"


def test_login_inside_one_clients_job_may_still_reuse_that_clients_password(capture_login):
    eng.CREDS.update(tin="1000000001", password="pw-A")
    with pytest.raises(eng.StopRequested):
        eng.login(object(), None, None)                            # re-login during A's own job
    assert capture_login == {"tin": "1000000001", "password": "pw-A"}


@pytest.mark.parametrize("login, problem_word", [
    ("1000000001", None),
    ("1000 000001", "digits"),
    ("1000/00001", "digits"),
    ("..\\1000000001", "digits"),
    ("1000000009", "already"),
])
def test_login_id_must_be_digits_and_not_a_duplicate(login, problem_word):
    others = [{"name": "Z", "login": "1000000009", "password": "p"}]
    problem = store.client_problem({"name": "N", "login": login, "password": "p"}, others)
    if problem_word is None:
        assert problem is None
    else:
        assert problem and problem_word in problem.lower()
