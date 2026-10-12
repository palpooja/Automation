"""
modules/test_login.py - 'Test login' button on the Clients dialog.

Opens the portal with the Ledger engine's own launch_browser()/login(),
lets you type the security code, reports whether the portal menus
appeared, logs out and closes.  Downloads nothing.
"""
import contextlib

from core.runner import PrintToLog, Stopped
from engines import ledger_engine as eng


def test_login(ctx, client):
    from playwright.sync_api import sync_playwright
    if not client.get("password"):
        return False, "no usable password is saved - open Clients > Edit and type it again"
    ctx.add_secret(client["password"])           # masked in the log (finding H6)
    eng.STOP_EVENT.clear()
    eng.PROMPT_HANDLER = lambda message, kind, defaults, poll: ctx.ask(message, kind, defaults, poll)
    with contextlib.redirect_stdout(PrintToLog(ctx)), contextlib.redirect_stderr(PrintToLog(ctx)):
        with sync_playwright() as p:
            browser = eng.launch_browser(p)
            try:
                context = browser.new_context(accept_downloads=True,
                                              viewport={"width": 1400, "height": 900})
                page = context.new_page()
                page.set_default_timeout(eng.NAV_TIMEOUT_MS)
                try:
                    eng.login(page, client["login"], client["password"])
                except eng.StopRequested:
                    raise Stopped()
                ok = eng.logged_in(page)
                detail = ("portal menus are visible" if ok
                          else "login form gone but portal menus not seen - check in the browser")
                try:                        # portal logout link, if shown (checked at once)
                    from core.session import press_logout_link_if_shown
                    press_logout_link_if_shown(page)
                except Exception:
                    pass
                return ok, detail
            finally:
                browser.close()
