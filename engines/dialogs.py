"""engines/dialogs.py - one answer policy for portal pop-up questions (finding H7).

The portal asks 'Are you sure you want to generate Tax Credit Certificate?' before
creating certificates.  Without a handler Playwright answers Cancel; the old code
answered OK to EVERY question for the rest of the session.  Now:

* an information box (alert - only an OK button) is acknowledged;
* the generate-certificate question is answered OK only while the program is
  pressing Generate on purpose (`with guard.expecting():`), and the logout
  question only while it is logging out (`with guard.expecting("logout"):`);
* any other question is answered Cancel and written to the log.

One guard per browser page, shared by the WHT and PAYE engines, so two engines
on the same page cannot answer the same question differently.
"""
import contextlib
import re
import weakref

GENERATE_CONFIRM = re.compile(r"generate\s+(the\s+)?tax\s+credit\s+certificate", re.I)
# live run 3 Oct 2026: 'You have requested to be logged out, do you want to proceed?'
LOGOUT_CONFIRM = re.compile(r"logged\s*out|log\s*out", re.I)
EXPECTED = {"generate": GENERATE_CONFIRM, "logout": LOGOUT_CONFIRM}
_GUARDS = weakref.WeakKeyDictionary()


class DialogGuard:
    def __init__(self, page):
        self.messages = []
        self.log = None
        self._expect = None
        page.on("dialog", self._on_dialog)

    @contextlib.contextmanager
    def expecting(self, kind="generate"):
        self._expect = EXPECTED[kind]
        try:
            yield self
        finally:
            self._expect = None

    def _on_dialog(self, dialog):
        msg = dialog.message or ""
        self.messages.append(msg)
        try:
            if dialog.type == "alert":
                dialog.accept()
            elif dialog.type == "confirm" and self._expect is not None and self._expect.search(msg):
                dialog.accept()
            else:
                if self.log:
                    self.log(f"  The portal asked '{msg}' - answered Cancel (not expected).")
                dialog.dismiss()
        except Exception:                 # the page closed or the box already went
            pass


def guard_for(page, log=None) -> DialogGuard:
    g = _GUARDS.get(page)
    if g is None:
        g = _GUARDS[page] = DialogGuard(page)
    if log is not None:
        g.log = log
    return g
