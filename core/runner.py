"""
core/runner.py - how a module plugs into the window, and how a run talks
back to the window from the worker thread.

THE MODULE CONTRACT
    Every file in modules/ defines one subclass of Module with:
        id        short key used in settings/history   e.g. "returns"
        title     text shown in the vertical menu      e.g. "Returns"
        build_panel(parent) -> tk.Frame   its own options page
        collect() -> dict                 read the page into plain values
                                          (raise ValueError with a message
                                          if something is not filled in)
        run(ctx, client, options)         do the work for ONE client
    The window never touches a module's widgets; the module never
    touches the window.  They only meet through RunContext.

THREADING RULE
    All browser work runs in ONE worker thread.  tkinter widgets may only
    be used from the main thread (CPython issue #77660), so the worker
    never calls a widget; it puts messages on ctx.ui_queue and the window
    drains that queue with root.after().  Questions to the person (type
    the security code, choose dates) go the same way and the answer comes
    back through ctx.reply_queue.
"""
import datetime
import queue
import threading
import time
from pathlib import Path


REVIEW_WORDS = ("fail", "error", "incomplete", "check")
QUIET_STATUSES = ("skipped", "already downloaded", "already on disk")   # a second try found it already done


def needs_review(status) -> bool:
    """One rule, used by the Review page, the summary workbook and the client
    state: an outcome a person must look at before the file is finalised."""
    low = str(status or "").lower()
    return any(w in low for w in REVIEW_WORDS)


class Stopped(BaseException):
    """Raised inside the worker when Stop is pressed.  BaseException so a
    bare 'except Exception' in engine code cannot swallow it."""


class BrowserClosed(Stopped):
    """The Chrome window was closed by hand during a run (live log 6 Oct 2026, 20:36).  Treated like
    Stop: the run ends cleanly instead of starting the next steps on a browser that is gone."""


def browser_gone(page=None, exc=None) -> bool:
    """True when the browser window has been closed: the error says so (Playwright's
    TargetClosedError, also when wrapped in another error), or the page reports it is closed."""
    e, depth = exc, 0
    while e is not None and depth < 6:
        if isinstance(e, BrowserClosed) or type(e).__name__ == "TargetClosedError" or \
                "Target page, context or browser has been closed" in str(e):
            return True
        e, depth = (e.__cause__ or e.__context__), depth + 1
    if page is not None:
        try:
            return bool(page.is_closed())
        except Exception:
            return False
    return False


class RunContext:
    def __init__(self, ui_queue: queue.Queue, reply_queue: queue.Queue,
                 stop_event: threading.Event, output_root: Path, log_path: Path):
        self.ui_queue = ui_queue
        self.reply_queue = reply_queue
        self.stop_event = stop_event
        self.output_root = Path(output_root)
        self.log_path = Path(log_path)
        self._fh = open(self.log_path, "a", encoding="utf-8")
        self.client_label = ""          # shown in the yellow bar
        self.login_timeout_s = 420
        self.items = []                 # every file produced / item failed (for the summary)
        self._client = self._module = ""
        self._login = ""
        self._secrets = set()           # passwords that must never reach a file (finding H6)
        self.voice_choice = None        # Settings > voice ('1', '4' or 'off'); set by the window
        self.after_client = []          # [(title, fn(page))] - work to do after the client's steps (10.14)
        self.final_try = False          # True during that end-of-client work: nothing is put off again (10.15)
        self._lock = threading.Lock()   # log file is written from the worker and the window

    def add_secret(self, secret):
        if secret and len(str(secret)) >= 2:
            self._secrets.add(str(secret))

    @property
    def secrets(self):
        return tuple(self._secrets)

    def scrub(self, text) -> str:
        """text with every known password replaced by ***."""
        text = str(text)
        for sec in self._secrets:
            text = text.replace(sec, "***")
        return text

    def set_scope(self, client_name, client_login, module_title):
        self._client, self._login, self._module = client_name, client_login, module_title

    def record(self, item, status, detail="", file=""):
        """status: 'saved' | 'skipped' | 'no records' | 'failed' (free text ok).
        During the end-of-client second try (final_try) an item this client already has is UPDATED, not
        listed twice (10.15): a failure fixed by the second try becomes '<new status> on second try'."""
        new = {"client": self._client, "login": self._login,
               "module": self._module, "item": self.scrub(item), "status": str(status),
               "detail": self.scrub(detail)[:300], "file": str(file) if file else ""}
        old = self._same_item(new["item"]) if self.final_try else None
        if old is None:
            self.items.append(new)
            return
        if new["status"].lower() in QUIET_STATUSES:
            return                                   # nothing new (already on disk)
        if needs_review(old["status"]) and not needs_review(new["status"]):
            new["status"] = f"{new['status']} on second try"
        elif needs_review(old["status"]) and needs_review(new["status"]):
            new["status"] = old["status"]
            new["detail"] = f"second try: {new['detail']}"[:300] if new["detail"] else old["detail"]
        old.update({k: new[k] for k in ("status", "detail", "file")})

    def _same_item(self, item):
        for it in reversed(self.items):
            if it["client"] == self._client and it["login"] == self._login and it["item"] == item:
                return it
        return None

    def resolve(self, item, status, detail=""):
        """Mark this client's earlier item (one that needed review) as dealt with, e.g. a whole tax type that
        failed in the first try and was done by the second.  Returns True when an item was changed."""
        old = self._same_item(self.scrub(item))
        if old is None or not needs_review(old["status"]):
            return False
        old["status"] = status
        if detail:
            old["detail"] = self.scrub(detail)[:300]
        return True

    # ---- messages to the window (never block) ----
    def log(self, text: str, kind: str = ""):
        """kind: '' normal, 'ok' green, 'err' red, 'warn' amber."""
        stamp = datetime.datetime.now().strftime("%H:%M:%S")
        line = f"[{stamp}] {self.scrub(text)}"
        try:
            with self._lock:
                self._fh.write(line + "\n")
                self._fh.flush()
        except (OSError, ValueError):       # log file closed or disk full: the window still shows it
            pass
        self.ui_queue.put(("log", (line, kind)))

    def status(self, text: str):
        self.ui_queue.put(("status", text))

    def progress(self, client_i, client_n, module_title, step=""):
        self.ui_queue.put(("progress", (client_i, client_n, module_title, step)))

    def check(self):
        if self.stop_event.is_set():
            raise Stopped()

    def sleep(self, seconds: float):
        """A pause that still honours Stop."""
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            self.check()
            time.sleep(min(0.2, end - time.monotonic()))

    # ---- questions to the person ----
    def ask(self, message: str, kind: str = "continue", defaults=("", ""),
            poll=None):
        """Show the yellow question bar and wait for an answer.
        kind: 'login' | 'dates' | 'continue'.  poll() may return True to
        answer automatically (e.g. the login form disappeared)."""
        while not self.reply_queue.empty():
            self.reply_queue.get_nowait()
        self.ui_queue.put(("ask", (message, kind, defaults, self.client_label)))
        self.log(">>> " + message)
        deadline = time.monotonic() + (self.login_timeout_s if kind == "login" else 10 ** 9)
        while True:
            if time.monotonic() > deadline:
                self.ui_queue.put(("hide_ask", None))
                self.log("No answer within the login time limit (Settings) - stopping.", "warn")
                return ("q", None, None) if kind == "dates" else "q"
            try:
                return self.reply_queue.get(timeout=0.5)
            except queue.Empty:
                pass
            if self.stop_event.is_set():
                self.ui_queue.put(("hide_ask", None))
                return ("q", None, None) if kind == "dates" else "q"
            if poll is not None:
                try:
                    if poll():
                        self.ui_queue.put(("hide_ask", None))
                        return "done_in_browser"
                except Exception:
                    pass

    # ---- a question that does NOT stop the work (10.14) ----
    def attention(self, message: str):
        """Say 'Attention needed.', show the yellow bar with Skip / Keep waiting, and return at once -
        the caller keeps working and calls poll_answer()."""
        while not self.reply_queue.empty():
            self.reply_queue.get_nowait()
        try:
            from core import voice
            if self.voice_choice:
                voice.say(voice.attention_phrase(), self.voice_choice)
        except Exception:
            pass
        self.ui_queue.put(("ask", (message, "skipwait", ("", ""), self.client_label)))
        self.log(">>> " + message, "warn")

    def poll_answer(self):
        """The person's answer to attention(), or None (never waits)."""
        try:
            return self.reply_queue.get_nowait()
        except queue.Empty:
            return None

    def hide_question(self):
        self.ui_queue.put(("hide_ask", None))

    def close(self):
        try:
            self._fh.close()
        except Exception:
            pass


class Module:
    id = "base"
    title = "Base"
    description = ""

    def build_panel(self, parent):
        raise NotImplementedError

    def collect(self) -> dict:
        return {}

    def run(self, ctx: RunContext, client: dict, options: dict) -> dict:
        """Return a small dict for the history row, e.g. {'files': 12}."""
        raise NotImplementedError


class PrintToLog:
    """File-like object: makes engine code that uses print() write to the
    suite log instead.  Used with contextlib.redirect_stdout."""
    def __init__(self, ctx: RunContext):
        self.ctx = ctx
        self._buf = ""

    def write(self, text):
        self._buf += text
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            if line.strip():
                low = line.lower()
                kind = ("err" if any(w in low for w in ("error", "failed", "could not"))
                        else "ok" if any(w in low for w in ("saved", "finished", "logged in"))
                        else "")
                self.ctx.log(line, kind)

    def flush(self):
        pass
