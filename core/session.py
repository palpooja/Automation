"""
core/session.py - one browser session per client.

    with open_session(ctx, client) as page:
        module_a.run(ctx, client, options, page)
        module_b.run(ctx, client, options, page)

Login is the Ledger engine's login(): it fills TIN + password, waits for
you to type the security code (yellow bar, or Log In in the browser),
refills on a wrong code.  Logout and browser close happen on exit, even
after an error or Stop.  Every module therefore shares ONE login.
"""
import contextlib
import time

from core.runner import BrowserClosed, PrintToLog, Stopped, browser_gone  # noqa: F401 (re-exported)
from engines import ledger_engine as eng


LOGOUT_SELECTORS = ("#logoutFormISSP input[value='Logout']", "#logoutFormISSP input[type='submit']",
                    "#logoutFormISSP input", "a:has-text('Logout')", "input[value='Logout']")
LOGOUT_WAIT_S = 8          # how long the portal may take to finish logging out


def press_logout_link_if_shown(page):
    """Press the first VISIBLE portal Logout link, in any frame (Test Login).  Universal rule:
    checked at once - the old code pressed the first match and waited 3 s when it was hidden."""
    for fr in page.frames:
        try:
            loc = fr.locator("#logoutFormISSP a, a:has-text('Logout')")
            n = loc.count()
        except Exception:
            continue
        for i in range(min(n, 10)):
            try:
                one = loc.nth(i)
                if one.is_visible():
                    one.click(timeout=3000)
                    return True
            except Exception:
                continue
    return False


def _logout_button_present(page):
    for fr in list(page.frames):
        for sel in LOGOUT_SELECTORS:
            try:
                if fr.locator(sel).count():
                    return True
            except Exception:
                continue
    return False


def logout(page, log):
    """Press Logout, answer the portal's 'do you want to proceed?' with OK, and report
    'Logged out' only when the Logout button has really gone (live run 3 Oct 2026:
    the question was answered Cancel and the session stayed open)."""
    from engines.dialogs import guard_for
    if browser_gone(page):
        log("Not logged out of the URA portal: the browser window was closed.")
        return
    guard = guard_for(page, log)
    for round_ in (1, 2):
        for fr in list(page.frames):
            for sel in LOGOUT_SELECTORS:
                try:
                    loc = fr.locator(sel).first
                    if not loc.count():
                        continue
                    with guard.expecting("logout"):
                        loc.evaluate("el => el.click()")
                        end = time.monotonic() + LOGOUT_WAIT_S
                        while time.monotonic() < end:
                            page.wait_for_timeout(250)
                            if not _logout_button_present(page):
                                log("Logged out of the URA portal.")
                                return True
                    log("Logout was pressed but the portal did not confirm it - the browser is being "
                        "closed anyway (the portal ends the session after its own time limit).", "warn")
                    return False
                except Exception:
                    continue
        if round_ == 1:
            # 10.17: the Logout button lives on the portal home page; a step that ended on another screen
            # (live 9 Oct 2026: the PAYE certificate form) left the URA session open server-side.
            log("The Logout button is not on this screen - opening the portal home page first.")
            try:
                if not eng.ensure_portal_home(page):
                    break
            except Exception as e:
                log(f"The portal home page did not open ({type(e).__name__}).", "warn")
                break
    log("Could not find the Logout button - the browser is being closed anyway.")
    return False


def go_home(ctx, page, client, before=""):
    """Bring the portal back to its home page (menus visible) before a module
    starts.  Each module leaves the portal on its own last screen - e.g.
    Returns ends on 'View return history' - and the next module looks for
    the home-page menus.  Uses the Ledgers engine's routine: navigate the
    portal frame to its home address (stays logged in); if that fails, go
    in through the public site's Portal link; re-login only if the session
    has really ended.  On failure a screenshot and the frame list are saved
    in the log folder."""
    # close any extra tabs a module left open (return copies etc.)
    for p_ in list(page.context.pages):
        if p_ is not page:
            try:
                p_.close()
            except Exception:
                pass
    ctx.log(f"Returning to the portal home page{' before ' + before if before else ''}...")
    out = PrintToLog(ctx)
    try:
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
            eng.ensure_portal_home(page, client["login"], client["password"])
        ctx.log("Portal home page is open.", "ok")
        return True
    except eng.StopRequested:
        raise Stopped()
    except Exception as e:
        shot = ctx.log_path.parent / f"home_failed_{client['login']}.png"
        try:
            page.screenshot(path=str(shot), full_page=True)
        except Exception:
            shot = None
        ctx.log(f"Could not get back to the portal home page: {str(e).splitlines()[0]}", "err")
        for i, f in enumerate(page.frames):
            try:
                ctx.log(f"   frame {i}: {f.url}", "warn")
            except Exception:
                pass
        if shot:
            ctx.log(f"   screenshot saved: {shot}", "warn")
        raise RuntimeError("could not return to the portal home page "
                           "(screenshot and frame list are in the log)")


@contextlib.contextmanager
def open_session(ctx, client):
    from playwright.sync_api import sync_playwright
    if not client.get("password"):
        # never fall back to anyone else's password (finding H2)
        raise RuntimeError(f"No usable password is saved for {client.get('name', '')} "
                           f"({client.get('login', '')}) - open Clients > Edit and type it again.")
    ctx.add_secret(client["password"])           # masked in every log line and summary (H6)
    eng.CREDS["tin"], eng.CREDS["password"] = client["login"], client["password"]
    eng.CURRENT_TIN = client["login"]
    try:
        eng.STOP_EVENT.clear()
        eng.PROMPT_HANDLER = lambda message, kind, defaults, poll: ctx.ask(message, kind, defaults, poll)
        out = PrintToLog(ctx)
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
            with sync_playwright() as p:
                browser = eng.launch_browser(p)
                page = None
                try:
                    context = browser.new_context(accept_downloads=True,
                                                  viewport={"width": 1400, "height": 900})
                    eng.speed_up(context)
                    page = context.new_page()
                    page.set_default_timeout(eng.NAV_TIMEOUT_MS)
                    eng.BROWSER_KEEPALIVE = lambda: page.wait_for_timeout(50)
                    ctx.log(f"Logging in as {client['name']} ({client['login']})...")
                    try:
                        eng.login(page, client["login"], client["password"])
                    except eng.StopRequested:
                        raise Stopped()
                    eng.dismiss_cookie_banner(page)
                    eng.dismiss_pre_login_popup(page)
                    yield page
                finally:
                    if page is not None:
                        try:
                            logout(page, ctx.log)
                        except Exception:
                            pass
                    try:
                        browser.close()
                    except Exception:
                        pass
    finally:
        eng.BROWSER_KEEPALIVE = None
        eng.CREDS["tin"] = eng.CREDS["password"] = None   # nothing carries over to the next client
