"""The window starts, shows warnings for unreadable passwords, and the Clients dialog works.
Needs a display (on Linux: xvfb-run -a python -m pytest); skipped otherwise."""
import base64
import json

import pytest


@pytest.fixture
def app(fake_data_dir, monkeypatch):
    tk = pytest.importorskip("tkinter")
    try:
        root = tk.Tk()
        root.destroy()
    except tk.TclError:
        pytest.skip("no display")
    from core import store
    from tkinter import messagebox
    shown = []
    for name in ("showwarning", "showinfo", "showerror"):
        monkeypatch.setattr(messagebox, name, lambda *a, _n=name, **k: shown.append((_n, a)))
    store.CLIENTS_FILE.write_text(json.dumps([
        {"name": "A", "login": "1000000001", "password": store.protect_secret("pw-A")},
        {"name": "B", "login": "1000000002",
         "password": base64.b64encode(b"\xff\xfe-undecryptable").decode()}]), encoding="utf-8")
    import app as app_module
    a = app_module.App()
    a.update()
    a.after(500)
    a.update()
    yield a, app_module, shown
    a.destroy()


def test_window_builds_and_warns_about_the_unreadable_password(app):
    a, _, shown = app
    a.after(450, lambda: None)
    for _ in range(10):
        a.update()
        a.after(60)
    log = a.logbox.get("1.0", "end")
    assert "must be typed again" in log
    assert [c["password"] for c in a.clients] == ["pw-A", ""]


def test_run_refuses_to_start_while_a_password_is_missing(app):
    a, _, shown = app
    a.start(all_ticked=True)
    assert any(kind == "showerror" and "Passwords needed" in args[0] for kind, args in shown)
    assert a.worker is None


def test_clients_dialog_lists_the_client_needing_a_password(app):
    a, app_module, _ = app
    d = app_module.ClientsDialog(a)
    a.update()
    rows = d.lb.get(0, "end")
    assert any("password must be typed again" in r for r in rows)
    d.destroy()


def test_certificates_page_starts_with_generate_tcc_unticked(app):
    a, app_module, _ = app
    cert = next(m for m in app_module.MODULES if m.id == "certificates")
    assert cert.gen_tcc.get() is False and cert.paye_gen.get() is False


def test_import_from_excel_adds_clients_and_saves_passwords_encrypted(app, tmp_path, monkeypatch):
    """Bulk upload (user, 6 Oct 2026) through the real Clients window."""
    a, app_module, shown = app
    from openpyxl import Workbook
    from core import store
    wb = Workbook(); ws = wb.active
    ws.append(["Client name", "Login ID (TIN)", "Password"])
    ws.append(["New Co Limited", 1000000077, "pw-new"])
    ws.append(["Bad Row", "12 34", "pw-bad"])
    path = tmp_path / "clients.xlsx"; wb.save(path)
    monkeypatch.setattr(app_module.filedialog, "askopenfilename", lambda **k: str(path))
    asked = []
    monkeypatch.setattr(app_module.messagebox, "askokcancel", lambda t, m, **k: asked.append(m) or True)
    d = app_module.ClientsDialog(a)
    d.import_excel()
    assert [c["login"] for c in a.clients][-1] == "1000000077" and "1000000077" in a.ticked_logins
    assert "row 3" in asked[0] and "pw-" not in asked[0], "problem rows listed, passwords never shown"
    saved = store.CLIENTS_FILE.read_text(encoding="utf-8")
    assert "1000000077" in saved and "pw-new" not in saved, "password stored encrypted, not as typed"
    d.destroy()


def test_excel_template_button_writes_the_template(app, tmp_path, monkeypatch):
    a, app_module, shown = app
    target = tmp_path / "Clients_template.xlsx"
    monkeypatch.setattr(app_module.filedialog, "asksaveasfilename", lambda **k: str(target))
    d = app_module.ClientsDialog(a)
    d.excel_template()
    assert target.exists()
    d.destroy()


def test_voice_choice_is_on_the_settings_page_and_saved(app):
    a, app_module, shown = app
    sp = a.settings_page
    assert sp.values()["voice"] == "1", "sample 1 by default"
    sp.voice.set("Sample 4")
    assert sp.values()["voice"] == "4"
    sp.voice.set("Off")
    assert sp.values()["voice"] == "off"


def _run_with_fakes(a, app_module, monkeypatch, close_after=None, quiet=False, run_factory=None):
    """The real run loop of App.start; browser sessions and module work are stand-ins."""
    import contextlib
    import time as _t
    from core import session as _session, voice as _voice
    from playwright._impl._errors import TargetClosedError
    said, ran = [], []

    class FakePage:
        closed = False

        def is_closed(self):
            return self.closed

    @contextlib.contextmanager
    def fake_session(ctx, client):
        yield FakePage()
    monkeypatch.setattr(_session, "open_session", fake_session)
    monkeypatch.setattr(_session, "go_home", lambda *a, **k: None)
    from core import net as _net                           # the run loop, not the internet, is under test
    monkeypatch.setattr(_net, "wait_online", lambda *a, **k: True)
    monkeypatch.setattr(_voice, "say", lambda parts, choice="1": said.append(parts) or True)

    def make_run(m):
        def run(ctx, client, options, page):
            ran.append((client["login"], m.title))
            if (client["login"], m.title) == close_after:
                page.closed = True                         # the person closes Chrome
                if not quiet:
                    raise TargetClosedError()
            return {"files": 1}
        return run
    for m in app_module.MODULES:
        monkeypatch.setattr(m, "run", (run_factory or make_run)(m))
    a.clients[1]["password"] = "pw-B"
    a.clients[1].pop("password_unreadable", None)
    a.ticked_logins = {c["login"] for c in a.clients}
    a.start(all_ticked=True)
    end = _t.monotonic() + 30
    while a.worker is not None and a.worker.is_alive() and _t.monotonic() < end:
        a.update()
        _t.sleep(0.05)
    settle = _t.monotonic() + 1.5                          # the window copies log lines on a timer
    while _t.monotonic() < settle:
        a.update()
        _t.sleep(0.05)
    return said, ran, a.logbox.get("1.0", "end")


def test_voice_says_starting_another_between_clients(app, monkeypatch):
    a, app_module, shown = app
    said, ran, log = _run_with_fakes(a, app_module, monkeypatch)
    assert said == [["Job completed.", "Starting another."], ["Job completed."]], said
    assert f"URA Suite {app_module.VERSION}" in log, "10.18: every run log starts with the build that made it"
    assert app_module.VERSION in a.title()


def test_closing_chrome_by_hand_stops_the_run_cleanly(app, monkeypatch):
    a, app_module, shown = app
    first = app_module.MODULES[0].title
    said, ran, log = _run_with_fakes(a, app_module, monkeypatch, close_after=("1000000001", first))
    assert ran == [("1000000001", first)], f"nothing more started after Chrome was closed: {ran}"
    assert "The browser window was closed" in log and "failed for" not in log, log[-1500:]
    assert said == [["Job stopped."]], said


def test_chrome_closed_during_the_last_step_does_not_start_the_next_client(app, monkeypatch):
    """The last step may finish without an error after the window is gone (month-by-month errors)."""
    a, app_module, shown = app
    last = app_module.MODULES[-1].title
    said, ran, log = _run_with_fakes(a, app_module, monkeypatch, close_after=("1000000001", last), quiet=True)
    assert all(login == "1000000001" for login, _ in ran), f"second client started: {ran}"
    assert "The browser window was closed" in log and said == [["Job stopped."]], (said, log[-800:])


def test_work_left_for_the_end_of_a_client_runs_after_its_other_steps(app, monkeypatch):
    """10.14: Returns sets a slow return aside; its second try runs after the client's other steps."""
    a, app_module, shown = app
    order = []
    first = app_module.MODULES[0]

    def factory(m):
        def run(ctx, client, options, page):
            order.append((client["login"], m.title))
            if m is first:
                ctx.after_client.append(("Returns - second try for 1 slow return(s)",
                                         lambda pg, lg=client["login"]: order.append((lg, "second try"))))
            return {"files": 1}
        return run
    said, ran, log = _run_with_fakes(a, app_module, monkeypatch, run_factory=factory)
    titles = [m.title for m in app_module.MODULES]
    for login in ("1000000001", "1000000002"):
        mine = [w for lg, w in order if lg == login]
        assert mine == titles + ["second try"], mine
    assert "----- Returns - second try for 1 slow return(s) -----" in log


def test_a_step_that_failed_as_a_whole_runs_again_after_the_other_steps(app, monkeypatch):
    """10.15 (user: 'Put Retry if anything else is failed'): a step that stopped with an error is run once
    more at the end of the client; its '(whole module)' line then says so."""
    a, app_module, shown = app
    first = app_module.MODULES[0]
    order, tries = [], {}

    def factory(m):
        def run(ctx, client, options, page):
            order.append((client["login"], m.title, ctx.final_try))
            key = (client["login"], m.title)
            tries[key] = tries.get(key, 0) + 1
            if m is first and client["login"] == "1000000001" and tries[key] == 1:
                raise RuntimeError("the portal page did not open")
            return {"files": 1}
        return run
    said, ran, log = _run_with_fakes(a, app_module, monkeypatch, run_factory=factory)
    mine = [(t, last) for lg, t, last in order if lg == "1000000001"]
    titles = [m.title for m in app_module.MODULES]
    assert mine == [(t, False) for t in titles] + [(first.title, True)], mine
    assert [t for lg, t, _ in order if lg == "1000000002"] == titles, "the other client: no second try"
    assert f"----- {first.title} - second try -----" in log
    assert f"{first.title}: second try finished." in log


def test_overview_counts_the_files_the_step_recorded_also_when_it_ends_with_an_error(app, monkeypatch, fake_data_dir):
    """10.17 (live 9 Oct 2026): Certificates saved 15 WHT files, then the PAYE part raised - the Overview showed
    0 files; Returns showed 41 in the Overview but 42 in the Files sheet (its status workbook).  The Overview now
    counts the file rows the step recorded, the same rows the Files sheet lists."""
    a, app_module, shown = app
    from core import store
    first, second = app_module.MODULES[0], app_module.MODULES[1]

    def factory(m):
        def run(ctx, client, options, page):
            if m is first:                              # 2 documents + a status workbook, reports 2
                ctx.record("A: return 1", "saved", "", "C:/out/r1.pdf")
                ctx.record("A: return 2", "saved", "", "C:/out/r2.pdf")
                ctx.record("Returns_Status.xlsx", "saved", "", "C:/out/Returns_Status.xlsx")
                return {"files": 2}
            if m is second:                             # files saved, then the step dies
                ctx.record("WHT Payee JUL-2025", "saved", "", "C:/out/w1.xlsx")
                ctx.record("WHT Payee AUG-2025", "saved", "", "C:/out/w2.xlsx")
                raise RuntimeError("the PAYE page did not open")
            return {"files": 1}
        return run
    said, ran, log = _run_with_fakes(a, app_module, monkeypatch, run_factory=factory)
    rows = {(h["login"], h["module"]): h for h in store.load_history() if h["login"] == "1000000001"}
    assert rows[("1000000001", first.title)]["files"] == 3, rows
    assert rows[("1000000001", second.title)]["files"] == 2, rows
    assert rows[("1000000001", second.title)]["status"].startswith("Error"), rows
    third = app_module.MODULES[2]
    assert rows[("1000000001", third.title)]["files"] == 1, "a step that records no file rows keeps its own count"
