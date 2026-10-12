"""
engines/wht_engine.py - drives 'Generate Tax Credit Certificate for
Withholding' month by month and captures the results grid.

RECONSTRUCTED from ura_bot.pyc (Python 3.13 bytecode) + ura_bot_fast.py.
Selectors (SEL), the calendar-scanning JavaScript (CAL_SCAN), the print
CSS, the keyword list, the loop in run_months() and every log message are
taken from the bytecode.  Control flow is a faithful translation; where
the bytecode was ambiguous the simpler reading was used and is marked
"(assumed)".

Differences from the original, on purpose:
  * WhtBot works on a Playwright `page` it is GIVEN, instead of starting
    its own browser.  That is what lets one login serve several modules.
  * Stop is checked through a callable (ctx.check), not a threading.Event.
"""
import base64
import re
import time

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]

LOGIN_URL = "https://ura.go.ug/en/etax-login/"
HOME_URL = "https://ura.go.ug/"

SEL = {
    "user": ["#userName", "input[name='userName']"],
    "password": ["#loginForm input[type='password']", "input[type='password']"],
    "captcha": ["#ans", "input[name='ans']"],
    "login_btn": ["#loginForm input[value='Log In']", "#loginForm input[type='submit']",
                  "#loginForm table tr:nth-of-type(6) input"],
    "login_back": ["#loginRespSec input", "#loginRespSec > input"],
    "logged_in_marker": ["#logoutFormISSP", "#mailbox", "#user-slide-menu"],
    "menu_item": "Generate Tax Credit Certificate for Withholding",
    "form_root": "#GenTaxCreditCert",
    "from_date": ["#frmDt", "input[name='frmDt']"],
    "date_img": ["#GenTaxCreditCert img[title*='Date' i]", "#GenTaxCreditCert img[src*='Calendar' i]",
                 "#GenTaxCreditCert img[onclick*='frmDt']", "#GenTaxCreditCert img[alt*='Date' i]",
                 "#GenTaxCreditCert table:nth-of-type(1) tr:nth-of-type(4) td:nth-of-type(2) img",
                 "#GenTaxCreditCert img"],
    "generate_for": ["#generateFor", "select[name='generateFor']"],
    "to_date": ["#toDt", "input[name='toDt']"],
    "search_btn": ["#GenTaxCreditCert input[value='Search']", "input[value='Search']",
                   "#GenTaxCreditCert tr.tdbar input:nth-of-type(1)"],
    "processing": ["text=Please wait as we process your request", "text=Your query is being processed",
                   "text=Please wait", "text=Processing"],
    "processing_close": ["#process-static-modal button.btn-secondary",
                         "button:has-text('Close')", "input[value='Close']", "a:has-text('Close')",
                         ".modal button.close"],
    # live recording 3 Oct 2026: Generate TCC -> summary screen with a Back button; no PDF
    "generate_btn": ["#generateTCCBtn", "input[value*='Generate TCC' i]", "button:has-text('Generate TCC')"],
    "tcc_summary": ["#TCCSummary"],
    "summary_back": ["#TCCSummary input[value='Back']", "#TCCSummary input[type=button]",
                     "#TCCSummary button:has-text('Back')"],
    "results_table": ["table.FORM_TAB_BORDER", "table:has(th:has-text('Agent TIN'))",
                      "table:has(td:has-text('Agent TIN'))"],
    "next_page": ["a:has-text('Next')", "a[title*='Next' i]", "input[value='Next']",
                  "a:has-text('>>')", "a.next", "li.next a"],
    "logout_btn": ["#logoutFormISSP input", "#logoutFormISSP > table input",
                   "a:has-text('Logout')", "input[value='Logout']"],
    "dismiss": ["#NINBRNMsg button", "#myModal button", "#parishMsgA button",
                "#parishMsgA h5 button", "div.modal button.close"],
}

CAL = {k: f"[data-bot='{k}']" for k in ("prevm", "nextm", "prevy", "nexty", "day1", "close")}

# Finds the open date-selector, tags its arrows and the '1' cell with
# data-bot attributes so Playwright can click them, and reads its heading.
CAL_SCAN = r"""
() => {
  const norm = s => (s || '').replace(/[\u00a0\s]+/g, ' ').trim();
  document.querySelectorAll('[data-bot]').forEach(e =>
      e.removeAttribute('data-bot'));

  const GLYPH = {
    '\u2039': 'prevm', '<': 'prevm', '\u2190': 'prevm',
    '\u203a': 'nextm', '>': 'nextm', '\u2192': 'nextm',
    '\u00ab': 'prevy', '<<': 'prevy',
    '\u00bb': 'nexty', '>>': 'nexty',
  };

  let anchor = null;
  for (const el of document.querySelectorAll('td, div, span, a, button')) {
    if (el.children.length > 1) continue;          // innermost node only
    const key = GLYPH[norm(el.textContent)];
    if (!key) continue;
    const box = el.getBoundingClientRect();
    if (!box.width || !box.height) continue;       // hidden
    if (!el.hasAttribute('data-bot')) el.setAttribute('data-bot', key);
    if (!anchor) anchor = el;
  }
  if (!anchor) return {found: false};

  const root = anchor.closest('table') || anchor.closest('div');
  if (!root) return {found: false};

  let title = '';
  const head = root.tHead || root;
  for (const cell of head.querySelectorAll('td, th, div, span')) {
    if (cell.hasAttribute('data-bot')) continue;
    const t = norm(cell.textContent);
    if (t.length >= 4 && /\d{4}/.test(t) && t.length < 40) { title = t; break; }
  }
  if (!title) title = norm(head.textContent).slice(0, 60);

  let day1 = false;
  const body = root.tBodies && root.tBodies.length ? root.tBodies[0] : root;
  for (const cell of body.querySelectorAll('td')) {
    if (norm(cell.textContent) !== '1') continue;
    const cls = (cell.className || '').toLowerCase();
    if (/other|empty|disabled|weekend-other|old|new/.test(cls)) continue;
    if (/(^|\s)(wn|week)(\s|$)/.test(cls)) continue;   // 'wk' column: week 1 is not day 1 (live calendar: class 'day wn')
    const box = cell.getBoundingClientRect();
    if (!box.width || !box.height) continue;
    cell.setAttribute('data-bot', 'day1');
    day1 = true;
    break;
  }
  let close = false;                               // the calendar's own close button (x)
  for (const cell of (root.tHead || root).querySelectorAll('td, th, div, span, a, button')) {
    if (cell.children.length > 1 || norm(cell.textContent) !== '\u00d7') continue;
    const box = cell.getBoundingClientRect();
    if (!box.width || !box.height) continue;
    if (!cell.hasAttribute('data-bot')) cell.setAttribute('data-bot', 'close');
    close = true;
    break;
  }
  return {found: true, title: title, day1: day1, close: close};
}
"""

# Presses an element the way a mouse button does - down, up, click - sent to the element itself.
# Why (live run 5 Oct 2026): URA's date selector is 'The DHTML Calendar' (Dynarch); its arrows and
# dates react only to the button going down and up on the same cell (calendar.js: _add_evs listens
# to mousedown, tableMouseUp calls cellClick).  A plain script click (el.click(), versions 10.2-10.5)
# sends only 'click' and is ignored; a mouse click at screen coordinates can land on whatever lies
# on top (the open calendar covered the Search button).  Sending all three events to the element
# behaves like a person's click for both kinds of calendar and does not depend on screen position.
PRESS_JS = r"""el => {
  const r = el.getBoundingClientRect();
  const o = {bubbles: true, cancelable: true, composed: true, view: window, button: 0,
             clientX: r.left + r.width / 2, clientY: r.top + r.height / 2};
  el.dispatchEvent(new MouseEvent('mousedown', Object.assign({buttons: 1}, o)));
  el.dispatchEvent(new MouseEvent('mouseup', Object.assign({buttons: 0}, o)));
  el.dispatchEvent(new MouseEvent('click', Object.assign({buttons: 0}, o)));
  return true;
}"""

# Last resort to close the calendar: hide the box around the arrows that CAL_SCAN tagged (the
# calendar shows itself again the next time its icon is pressed).
CAL_HIDE_JS = r"""() => {
  const arrow = document.querySelector("[data-bot='prevm'], [data-bot='nextm'], [data-bot='prevy'], [data-bot='nexty']");
  if (!arrow) return false;
  const table = arrow.closest('table');
  if (!table) return false;
  let box = table;
  for (let p = table.parentElement; p && p !== document.body; p = p.parentElement) {
    const pos = getComputedStyle(p).position;
    if (pos === 'absolute' || pos === 'fixed') { box = p; break; }
  }
  box.style.display = 'none';
  return true;
}"""

DEAD_PRESSES = 2            # an arrow that leaves the heading unchanged this many times in a row is not working

PRINT_CSS = """
@page { size: A4 landscape; margin: 8mm; }
header, nav, footer, .site-header, .site-footer, .navbar, .breadcrumb,
[id*='chat' i], [class*='chat' i], [id*='cookie' i], iframe:not(#__bot_main) {
  display: none !important;
}
#__bot_main { width: 100% !important; border: 0 !important; }
"""

KEYWORDS = ("tin", "agent", "amount", "withheld", "gross", "period", "creditable", "tcc",
            "nature", "payment", "return", "name", "sr.no", "srno", "assessment", "payee",
            "payer", "certificate")
RELOAD_GRACE_S = 6                 # after the portal reloads: wait this long for a table (x speed setting)
RESULT_MATCH = "TIN"            # pandas.read_html(match=...) (assumed: 'Agent TIN' column)
POLL_MS = 150
RESULT_WAIT_MS = 60000          # how long a search may take before the month is FAILED
ONLINE_WAIT = None              # suite: waits while the internet is down

NO_RECORD_RE = r"no\s+(record|data|result|detail)s?\s*(found|available)?\b"

MARK_STALE_JS = r"""() => {
  document.documentElement.setAttribute('data-bot-stale', '1');      // gone after a page reload
  document.querySelectorAll('table, table *').forEach(e => e.setAttribute('data-bot-stale', '1'));
  document.querySelectorAll('body *').forEach(e => {
    if (e.children.length === 0 &&
        /no\s+(record|data|result|detail)s?/i.test(e.textContent || ''))
      e.setAttribute('data-bot-stale', '1');
  });
}"""

# true when NEW results are on screen: a non-stale table that looks like the
# results grid, or a non-stale 'no records' message
# A results grid: has a TIN column heading, at least 5 columns, and does not
# hold the search form's own fields (date boxes, Generate For, Search) - the
# live portal run of 3 Oct 2026 showed the form being read as "7 rows x 4
# columns" of data).
IS_GRID_JS = r"""(t) => {
  if (t.querySelector('#frmDt, #toDt, #generateFor, [name=frmDt], [name=toDt], [name=generateFor], input[type=button][value=Search], input[type=submit][value=Search], input[type=button][value=Reset]')) return false;          // the search form itself
  const rows = t.querySelectorAll('tr');
  if (rows.length < 2) return false;
  let maxCols = 0;
  rows.forEach(r => { maxCols = Math.max(maxCols, r.children.length); });
  if (maxCols < 5) return false;
  const head = (rows[0].innerText || '').toLowerCase();
  return /\btin\b/.test(head) || head.includes('agent');
}"""

FRESH_JS = r"""() => {
  const vis = e => { const b = e.getBoundingClientRect(); return b.width > 0 && b.height > 0; };
  const isGrid = """ + IS_GRID_JS + r""";
  for (const t of document.querySelectorAll('table')) {
    const fresh = !t.hasAttribute('data-bot-stale') || t.querySelector('tr:not([data-bot-stale])');
    if (fresh && vis(t) && isGrid(t)) return 'table';
  }
  for (const e of document.querySelectorAll('body *:not([data-bot-stale])')) {
    if (e.children.length === 0 && vis(e) &&
        /no\s+(record|data|result|detail|transaction|certificate)s?|not\s+found/i.test(e.textContent || ''))
      return 'none';
  }
  if (!document.documentElement.hasAttribute('data-bot-stale') && document.readyState === 'complete'
      && document.querySelector('#frmDt, [name=frmDt], #generateFor'))
    return 'reloaded';     // the WHT page itself came back (8 Oct 2026: not any other frame that reloaded)
  return '';
}"""

CLEAN_HTML_JS = """() => {
  const c = document.documentElement.cloneNode(true);
  // remove an old element only when nothing new has appeared inside it
  c.querySelectorAll('[data-bot-stale]').forEach(e => {
    if (e !== c && e.isConnected && !e.querySelector(':not([data-bot-stale])')) e.remove();
  });
  // a value shown in a text box inside a results cell (e.g. the TCC number in
  // #tccNo1) is copied in as plain text, or the Excel column would be empty
  c.querySelectorAll('table input[type=text], table input:not([type])').forEach(i => {
    if (!i.closest('table').querySelector('#frmDt, #toDt, #generateFor'))
      i.replaceWith(document.createTextNode(i.value || i.getAttribute('value') || ''));
  });
  // tables holding form controls are the search form, never results
  c.querySelectorAll('table').forEach(t => {
    if (t.isConnected && t.querySelector('#frmDt, #toDt, #generateFor, [name=frmDt], [name=toDt], [name=generateFor], input[type=button][value=Search], input[type=submit][value=Search], input[type=button][value=Reset]')) t.remove();
  });
  return c.outerHTML;
}"""


class Cancelled(BaseException):
    pass


class SetupError(Exception):
    """A problem that will be the same for every month (e.g. the portal has no
    such 'Generate For' option): stop that type at once instead of retrying
    each month."""


# Tick exactly the result rows that have no TCC number yet (one script click per
# box, only when its state differs) and report what is on screen.
TCC_TICK_JS = r"""(doTick) => {
  const isGrid = """ + IS_GRID_JS + r""";
  const norm = s => (s || '').replace(/[\u00a0\s]+/g, ' ').trim();
  for (const t of document.querySelectorAll('table')) {
    if (!isGrid(t)) continue;
    const rows = [...t.querySelectorAll('tr')];
    const head = rows[0];
    const heads = [...head.children].map(c => norm(c.innerText).toLowerCase());
    const col = heads.findIndex(h => h.includes('tcc'));
    let total = 0, missing = 0, ticked = 0;
    for (const r of rows.slice(1)) {
      const cells = [...r.children];
      if (cells.length < 3) continue;
      total++;
      let tcc = '';
      const box = r.querySelector('[id^=tccNo]') || (col >= 0 ? cells[col] : null);
      if (box) {
        const inp = box.matches('input') ? box : box.querySelector('input');
        tcc = norm(inp ? inp.value : box.innerText);
      }
      const cb = r.querySelector('input[type=checkbox]');
      const want = !tcc;
      if (want) missing++;
      if (doTick && cb && !cb.disabled) {
        if (cb.checked !== want) cb.click();
        if (cb.checked && want) ticked++;
      }
    }
    return {found: true, tccCol: col, total, missing, ticked};
  }
  return {found: false};
}"""


class WhtBot:
    def __init__(self, page, log, check, screenshot_dir=None):
        from engines.dialogs import guard_for
        self._guard = guard_for(page, log)       # one answer policy per page (finding H7)
        self.dialogs = self._guard.messages
        self.page = page
        self.log = log
        self.check = check                 # raises when Stop is pressed
        self.screenshot_dir = screenshot_dir
        self._cdp = None
        self.click_log = []
        self.expected_from = None          # the From date pick_month set; search() checks it is still there

    # ------------------------------------------------------------ frames / find
    def frames(self):
        try:
            return list(self.page.frames)
        except Exception:
            return [self.page]

    def describe_frames(self):
        out = []
        for fr in self.frames():
            try:
                out.append(f"{fr.name or '(main)'}: {fr.url[:80]}")
            except Exception:
                pass
        return "; ".join(out)

    def find(self, sels, frame=None, timeout=3000):
        """One look per frame; waits only for an existing element to be visible."""
        if isinstance(sels, str):
            sels = [sels]
        for fr in ([frame] if frame else self.frames()):
            for sel in sels:
                try:
                    loc = fr.locator(sel).first
                    if loc.count() == 0:
                        continue
                    loc.wait_for(state="visible", timeout=timeout)
                    return loc
                except Exception:
                    continue
        return None

    def wait_find(self, sels, timeout=10000, frame=None, visible=True):
        """Keeps looking (every 0.15 s) until the element shows up."""
        if isinstance(sels, str):
            sels = [sels]
        end = time.monotonic() + timeout / 1000.0
        while True:
            self.check()
            for fr in ([frame] if frame else self.frames()):
                for sel in sels:
                    try:
                        loc = fr.locator(sel).first
                        if loc.count() and (not visible or loc.is_visible()):
                            return loc
                    except Exception:
                        continue
            if time.monotonic() >= end:
                return None
            self.page.wait_for_timeout(POLL_MS)

    def click_any(self, sels, label, frame=None, timeout=5000, required=True, force=True):
        """Normal click, then (if force) a forced click at the element's screen position, then a
        script click.  force=False for buttons where something may lie on top (live 5 Oct 2026:
        the forced Search click landed on the open calendar and changed the From date)."""
        self.check()
        loc = self.find(sels, frame=frame, timeout=timeout)
        if loc is None:
            if required:
                raise RuntimeError(f"Could not find: {label}")
            self.log(f"  (skipped - {label} not on the page)")
            return False
        try:
            loc.scroll_into_view_if_needed(timeout=3000)
        except Exception:
            pass
        last = None
        for how in (("normal", "force", "script") if force else ("normal", "script")):
            try:
                if how == "normal":
                    loc.click(timeout=timeout)
                elif how == "force":
                    loc.click(timeout=4000, force=True)
                else:
                    loc.evaluate("el => el.click()")
                    self.log(f"  (used a scripted click for {label})")
                self.click_log.append((label, how))
                return True
            except Exception as exc:
                last = exc
        if required:
            raise RuntimeError(f"Could not click {label}: {last}")
        return False

    def settle(self, ms=800):
        try:
            self.page.wait_for_load_state("load", timeout=30000)
        except Exception:
            pass
        self.page.wait_for_timeout(ms)

    def snapshot(self, name):
        if not self.screenshot_dir or not self.page:
            return None
        try:
            import os
            path = os.path.join(self.screenshot_dir, f"screen_{name}.png")
            self.page.screenshot(path=path, full_page=True)
            return path
        except Exception:
            return None

    # ------------------------------------------------------------ login state
    def is_logged_in(self):
        sel = SEL["logged_in_marker"] + SEL["from_date"]
        for fr in self.frames():
            for s in sel:
                try:
                    if fr.locator(s).count() > 0:
                        return True
                except Exception:
                    continue
        return False

    # ------------------------------------------------------------ pop-ups
    _POPUP_NAMES = ("Skip", "Close", "OK", "Continue")

    def dismiss_popups(self, rounds=3):
        """The 'Skip' notices sit inside the portal iframe.  The second one
        only appears once the first is closed, hence the rounds."""
        for rnd in range(rounds):
            hit = False
            for fr in self.frames():
                for sel in SEL["dismiss"]:
                    try:
                        loc = fr.locator(sel).first
                        if loc.count() and loc.is_visible():
                            try:
                                loc.click(timeout=3000)
                            except Exception:
                                loc.evaluate("el => el.click()")
                            self.log(f"Dismissed a pop-up ({sel}).")
                            hit = True
                            self.page.wait_for_timeout(800)
                    except Exception:
                        continue
                for name in self._POPUP_NAMES:
                    try:
                        btn = fr.get_by_role("button", name=name, exact=True).first
                        if btn.count() and btn.is_visible():
                            btn.click(timeout=2500)
                            self.log(f"Dismissed a pop-up ('{name}').")
                            hit = True
                            self.page.wait_for_timeout(800)
                    except Exception:
                        continue
            if not hit:
                if rnd == 0:
                    self.log("No pop-ups to dismiss.")
                return

    def _quick_dismiss(self):
        hit = False
        for fr in self.frames():
            targets = [fr.locator(sel).first for sel in SEL["dismiss"]]
            targets += [fr.get_by_role("button", name=n, exact=True).first
                        for n in self._POPUP_NAMES]
            for loc in targets:
                try:
                    if loc.count() and loc.is_visible():
                        try:
                            loc.click(timeout=2000)
                        except Exception:
                            loc.evaluate("el => el.click()")
                        self.log("Dismissed a pop-up.")
                        hit = True
                        self.page.wait_for_timeout(200)
                except Exception:
                    continue
        return hit

    # ------------------------------------------------------------ the WHT page
    @staticmethod
    def _absolute(fr, href):
        try:
            return fr.evaluate("u => new URL(u, document.baseURI).href", href)
        except Exception:
            return href

    def _menu_link(self):
        label = SEL["menu_item"]
        for fr in self.frames():
            for make in (lambda: fr.get_by_role("link", name=label, exact=False),
                         lambda: fr.locator(f"a:has-text({label!r})")):
                try:
                    link = make().first
                    if link.count():
                        return fr, link
                except Exception:
                    continue
        return None, None

    def open_wht_page(self, quiet=False):
        """Quick route from ura_bot_fast.py: script-click the menu link (it may
        be in a closed sub-menu), then wait for the form; plan B follows the
        link's address inside its frame; last resort hovers the parent."""
        self.check()
        if not quiet:
            self.log("Opening 'Generate Tax Credit Certificate for Withholding'...")
        if self.wait_find(SEL["from_date"], timeout=0) is not None:
            if not quiet:
                self.log("Search form is open.")
            return
        end = time.monotonic() + 40
        tried_href = False
        while time.monotonic() < end:
            self.check()
            self._quick_dismiss()
            fr, link = self._menu_link()
            if link is None:
                self.page.wait_for_timeout(POLL_MS)
                continue
            try:
                href = (link.get_attribute("href") or "").strip()
            except Exception:
                href = ""
            try:
                link.evaluate("el => el.click()")
                self.click_log.append((SEL["menu_item"], "script"))
            except Exception:
                self.page.wait_for_timeout(POLL_MS)
                continue
            if self.wait_find(SEL["from_date"], timeout=12000) is not None:
                if not quiet:
                    self.log("Search form is open.")
                return
            if href and not tried_href:
                tried_href = True
                try:
                    if href.lower().startswith("javascript:"):
                        fr.evaluate("code => eval(code)", href.split(":", 1)[1])
                    else:
                        fr.evaluate("u => { window.location.href = u; }",
                                    self._absolute(fr, href))
                    if self.wait_find(SEL["from_date"], timeout=12000) is not None:
                        if not quiet:
                            self.log("Search form is open.")
                        return
                except Cancelled:
                    raise
                except Exception:
                    pass
        # slow route: hover the parent menu, then click
        self.log("  quick route did not open the page - trying the slow route...")
        fr, link = self._menu_link()
        if link is not None:
            try:
                parent = link.locator("xpath=ancestor::li[1]/..").first
                parent.hover(timeout=3000)
                self.page.wait_for_timeout(500)
                link.click(timeout=5000)
            except Exception:
                pass
            if self.wait_find(SEL["from_date"], timeout=15000) is not None:
                self.log("Search form is open.")
                return
        raise RuntimeError("Could not open the WHT search form. Frames: " + self.describe_frames())

    def form_frame(self):
        for fr in self.frames():
            for sel in SEL["from_date"]:
                try:
                    if fr.locator(sel).count():
                        return fr
                except Exception:
                    continue
        raise RuntimeError("The search form did not load.")

    # ------------------------------------------------------------ calendar
    def _poll(self, read, ok, max_ms, step_ms=100):
        """Call read() until ok(result) or max_ms has passed, then return the last
        result.  Same final check at the same moment as the fixed pause it
        replaces; it only stops early when the check already succeeds."""
        end = time.monotonic() + max_ms / 1000.0
        while True:
            value = read()
            if ok(value) or time.monotonic() >= end:
                return value
            self.page.wait_for_timeout(min(step_ms, max(1, int((end - time.monotonic()) * 1000))))

    def _open_calendar(self, frame):
        """Open the date selector.  A script click first (the live portal took 6-11 s to
        accept a mouse click here), confirmed by reading the calendar's heading; only if
        that did not open it, the old click sequence (normal, forced, script)."""
        def is_open(s):
            return bool(s.get("found")) and self._parse_title(s.get("title")) is not None
        loc = self.find(SEL["date_img"], frame=frame, timeout=5000)
        if loc is not None:
            try:
                loc.evaluate("el => el.click()")
                scan = self._poll(lambda: self._cal_scan(frame), is_open, 3000)
                if is_open(scan):
                    self.click_log.append(("the Date Selector icon", "script"))
                    return scan
            except Exception:
                pass
        self.click_any(SEL["date_img"], "the Date Selector icon", frame=frame)
        return self._poll(lambda: self._cal_scan(frame), is_open, 600)

    def _cal_scan(self, frame):
        try:
            return frame.evaluate(CAL_SCAN) or {"found": False}
        except Exception:
            return {"found": False}

    @staticmethod
    def _parse_title(text):
        """'June 2026' / 'Jun, 2026' -> (2026, 6); None if unreadable."""
        if not text:
            return None
        year = re.search(r"(19|20)\d{2}", text)
        if not year:
            return None
        low = text.lower()
        for i, name in enumerate(MONTHS, start=1):
            if name.lower()[:3] in low:
                return int(year.group(0)), i
        return None

    @staticmethod
    def _cal_open(scan):
        return bool(scan.get("found")) and WhtBot._parse_title(scan.get("title")) is not None

    def _press(self, locator):
        """Press like a mouse button (down, up, click) on the element itself - see PRESS_JS."""
        locator.evaluate(PRESS_JS)

    def close_calendar(self, frame):
        """Make sure no date selector is showing.  On 5 Oct 2026 an open calendar lay over the
        Search button.  Its own close button (x) first, then hide its box.  True when closed."""
        scan = self._cal_scan(frame)
        if not self._cal_open(scan):
            return True
        if scan.get("close"):
            try:
                self._press(frame.locator(CAL["close"]).first)
            except Exception:
                pass
            scan = self._poll(lambda: self._cal_scan(frame), lambda s: not self._cal_open(s), 1500)
            if not self._cal_open(scan):
                return True
        try:
            frame.evaluate(CAL_HIDE_JS)
        except Exception:
            pass
        if self._cal_open(self._cal_scan(frame)):
            self.log("  WARNING: the date selector could not be closed.")
            return False
        self.log("  (closed the date selector by hiding it)")
        return True

    def _direct_set(self, frame, value):
        """Close the calendar FIRST, then write the date: a calendar left open overwrote the
        written date on 5 Oct 2026 (every wrong From date was a Tuesday - the cell under Search)."""
        self.close_calendar(frame)
        self._js_set_date(frame, value)

    def pick_month(self, frame, year, month):
        """Open the date selector, walk to the wanted month, press day 1, close the selector,
        then make sure the portal's 'To' date followed.  Returns a list of warning strings
        (empty when everything went cleanly)."""
        self.check()
        warnings = []
        target = f"01/{month:02d}/{year}"
        self.expected_from = target
        self.log(f"Selecting {MONTHS[month - 1]} {year} from the calendar...")
        self.clear_overlays(frame)
        scan = self._open_calendar(frame)
        if not scan.get("found"):
            self.log("  Calendar did not open - falling back to a direct set.")
            warnings.append("calendar did not open; date written directly")
            self._direct_set(frame, target)
            warnings += self.sync_to_date(frame, year, month)
            return warnings
        start_title = scan.get("title", "")
        self.log(f"  Calendar opened on '{start_title}'.")
        start = self._parse_title(start_title)
        presses = 0

        def step(key):
            # mouse-like press (PRESS_JS); the heading is read back after every press
            nonlocal presses
            before = scan.get("title")
            self._press(frame.locator(CAL[key]).first)
            presses += 1
            return self._poll(lambda: self._cal_scan(frame),
                              lambda s: s.get("title") != before, 1500)

        # years first (if the calendar has year arrows), then months.  A press that leaves the
        # heading unchanged DEAD_PRESSES times in a row means the arrows are not working: stop
        # (5 Oct 2026: 60 dead presses = about 110 s per month, then a wrong date anyway).
        dead, stuck = 0, False
        for phase, limit in (("years", 12), ("months", 60)):
            for _ in range(limit):
                self.check()
                cur = self._parse_title(scan.get("title"))
                if cur is None:
                    self.log(f"  Could not read the heading ('{scan.get('title', '')}').")
                    stuck = True
                    break
                if phase == "years":
                    if cur[0] == year:
                        break
                    key = "prevy" if cur[0] > year else "nexty"
                    try:
                        if frame.locator(CAL[key]).count() == 0:
                            break
                    except Exception:
                        break
                else:
                    if cur == (year, month):
                        break
                    key = "prevm" if cur > (year, month) else "nextm"
                new = step(key)
                if new.get("title") == scan.get("title"):
                    dead += 1
                    if dead >= DEAD_PRESSES:
                        stuck = True
                        break
                else:
                    dead = 0
                scan = new
            if stuck:
                break
        end = self._parse_title(scan.get("title"))
        if start and end and end != start:
            moved = abs((start[0] * 12 + start[1]) - (end[0] * 12 + end[1]))
            self.log(f"  Moved {moved} month(s) to '{scan.get('title', '')}' ({presses} press(es)).")
        if end != (year, month):
            if stuck and dead >= DEAD_PRESSES:
                self.log(f"  The calendar's arrow did not move after {DEAD_PRESSES} presses - "
                         "closing it and writing the date directly.")
                warnings.append("calendar arrows did not move; date written directly")
            else:
                self.log(f"  The calendar did not reach {MONTHS[month - 1]} {year} - "
                         "closing it and writing the date directly.")
                warnings.append("calendar did not reach the month; date written directly")
            self._direct_set(frame, target)
            warnings += self.sync_to_date(frame, year, month)
            return warnings
        if not scan.get("day1"):
            scan = self._cal_scan(frame)
        if not scan.get("day1"):
            self.log("  Could not find day 1 - falling back to a direct set.")
            warnings.append("day 1 not found; date written directly")
            self._direct_set(frame, target)
        else:
            self._press(frame.locator(CAL["day1"]).first)
            got = self._poll(lambda: self._read_date(frame),
                             lambda g: bool(g) and g.replace("-", "/") == target, 1500)
            if got and got.replace("-", "/") == target:
                self.log(f"  Period set to {got}.")
                self.close_calendar(frame)       # single-click calendars close themselves; make sure
            else:
                self.log(f"  Calendar produced '{got}', expected '{target}'. Correcting.")
                warnings.append(f"calendar gave '{got}'; date written directly")
                self._direct_set(frame, target)
        warnings += self.sync_to_date(frame, year, month)
        return warnings

    def sync_to_date(self, frame, year, month):
        """The portal fills 'Return Date To' from an onblur handler on the
        From box.  A calendar click never blurs it, so fire the handler
        ourselves and confirm the To month matches.  Searching with a stale
        To date is how a run silently returns the wrong month."""
        want = f"/{month:02d}/{year}"
        for attempt in range(4):
            to = self._read_to_date(frame)
            if to and to.replace("-", "/").endswith(want):
                self.log(f"  To date confirmed: {to}.")
                return []
            try:
                frame.evaluate("""() => {
                    const el = document.getElementById('frmDt')
                             || document.querySelector("input[name='frmDt']");
                    if (!el) return;
                    if (typeof getToDate === 'function') { try { getToDate(el); } catch (e) {} }
                    el.dispatchEvent(new Event('change', {bubbles: true}));
                    el.dispatchEvent(new Event('blur',   {bubbles: true}));
                }""")
            except Exception:
                pass
            self._poll(lambda: self._read_to_date(frame),          # was: fixed 600/1000/1400/1800 ms
                       lambda t: bool(t) and t.replace("-", "/").endswith(want), 600 + 400 * attempt)
        to = self._read_to_date(frame)
        if not to:
            self.log("  WARNING: the To date is blank - the portal may reject the search.")
            return ["To date blank"]
        self.log(f"  WARNING: To date is '{to}', not in {MONTHS[month - 1]} {year}. "
                 "Check the results are for the right month.")
        return [f"To date '{to}' did not follow the From date"]

    def _read_to_date(self, frame):
        for sel in SEL["to_date"]:
            try:
                if frame.locator(sel).count():
                    return frame.locator(sel).first.input_value() or ""
            except Exception:
                continue
        return ""

    def _read_date(self, frame):
        for sel in SEL["from_date"]:
            try:
                if frame.locator(sel).count():
                    return frame.locator(sel).first.input_value() or ""
            except Exception:
                continue
        return ""

    def _js_set_date(self, frame, value):
        for sel in SEL["from_date"]:
            try:
                if not frame.locator(sel).count():
                    continue
                frame.eval_on_selector(sel, """(el, v) => {
                    el.value = v;
                    el.dispatchEvent(new Event('input',  {bubbles: true}));
                    el.dispatchEvent(new Event('change', {bubbles: true}));
                }""", value)
                self.log(f"  WARNING: period written directly ({value}).")
                return False
            except Exception:
                continue
        raise RuntimeError("Could not set the period at all.")

    # Live portal, 3 Oct 2026: [['', ''], ['D', 'Payee ( Deductee )'], ['W', 'Withholding Agent']]
    GEN_WORDS = {"D": ("payee", "deductee"), "W": ("withholding agent", "payer", "deductor")}

    def set_generate_for(self, frame, code="D"):
        """D = Payee (Deductee), W = Withholding Agent (Payer).  Chooses by value,
        else by wording, and reads the choice back.  A missing option is a
        SetupError: it would fail the same way for every month."""
        words = self.GEN_WORDS.get(code, (code.lower(),))
        for sel in SEL["generate_for"]:
            dd = frame.locator(sel).first
            try:
                if not dd.count():
                    continue
                opts = dd.evaluate("el => Array.from(el.options).map(o => [o.value, o.text.trim()])",
                                   timeout=5000)
            except Exception:
                continue
            match = [v for v, t in opts if v == code] or \
                    [v for v, t in opts if any(w in t.lower() for w in words)]
            if not match:
                raise SetupError(f"the portal's 'Generate For' list has no option for "
                                 f"{'Payee' if code == 'D' else 'Withholding Agent (Payer)'}: "
                                 f"{[t for v, t in opts if t]}")
            dd.select_option(match[0], timeout=5000)
            got = dd.evaluate("el => el.options[el.selectedIndex].text", timeout=5000) or ""
            if not any(w in got.lower() for w in words):
                raise RuntimeError(f"the portal shows '{got.strip()}' after choosing")
            self.log(f"Generate for: {got.strip()}")
            return
        raise RuntimeError("the 'Generate For' box was not found on the form")

    # ------------------------------------------------------------ search
    def clear_overlays(self, frame=None, patience=40):
        """Wait out the 'Processing' box; close it if it stays."""
        seen = False
        for _ in range(patience):
            self.check()
            busy = self.find(SEL["processing"], frame=frame, timeout=250)
            if busy is None:
                return seen
            if not seen:
                self.log("  waiting for 'Processing' to finish...")
                seen = True
            self.page.wait_for_timeout(1000)
        for _ in range(3):
            if self.click_any(SEL["processing_close"], "the Processing Close button",
                              frame=frame, timeout=2500, required=False):
                self.log("  closed the 'Processing' box.")
                self.page.wait_for_timeout(600)
                if self.find(SEL["processing"], frame=frame, timeout=400) is None:
                    return True
        return True

    def search(self, frame):
        """Click Search and wait for NEW results.  Old results are marked
        stale first, so a slow portal can never make last month's table (or
        an empty screen) count as this month's answer.  No new table and no
        new 'no records' message within RESULT_WAIT_MS -> RuntimeError, which
        run_months records as a FAILED month (never as 'no records')."""
        self.check()
        self.log("Searching...")
        self.clear_overlays(frame)
        self.close_calendar(frame)                   # an open calendar covered Search on 5 Oct 2026
        before_from = self._read_date(frame)
        want = self.expected_from
        if want and before_from.replace("-", "/") != want:
            raise RuntimeError(f"the From date on screen is '{before_from}', not '{want}' - "
                               "Search was not pressed")
        for fr in self.frames():
            try:
                fr.evaluate(MARK_STALE_JS)
            except Exception:
                pass
        self.no_table_after_reload = False
        # never a forced click here: it presses the screen position, whatever lies on top
        self.click_any(SEL["search_btn"], "the Search button", frame=frame, force=False)
        end = time.monotonic() + RESULT_WAIT_MS / 1000.0
        noted = False
        while time.monotonic() < end:
            self.check()
            now_from = self._from_on_same_page(frame)
            if now_from is not None and before_from and now_from != before_from:
                raise RuntimeError(f"the From date changed from '{before_from}' to '{now_from}' when "
                                   "Search was pressed - the search did not run")
            for fr in self.frames():
                try:
                    kind = fr.evaluate(FRESH_JS)
                except Exception:
                    kind = ""
                if kind == "reloaded":
                    # the portal answered with a fresh page; give a grid time to draw
                    end2 = time.monotonic() + RELOAD_GRACE_S
                    while time.monotonic() < end2:
                        self.page.wait_for_timeout(300)
                        try:
                            again = fr.evaluate(FRESH_JS)
                        except Exception:
                            again = ""
                        if again in ("table", "none"):
                            return again
                    self.log("  The portal answered without a results table - "
                             "treated as no records (page copy saved for checking).")
                    self.no_table_after_reload = True
                    return "none"
                if kind:
                    self.page.wait_for_timeout(300)        # let the grid finish drawing
                    return kind
            if not noted and time.monotonic() > end - RESULT_WAIT_MS / 1000.0 + 10:
                self.log("  the portal is slow - still waiting for the results...")
                noted = True
            self.page.wait_for_timeout(250)
        raise RuntimeError(f"no results and no 'no records' message within "
                           f"{RESULT_WAIT_MS // 1000} s after Search - the portal may be slow "
                           f"(From date on screen: '{self._read_date(frame)}')")

    def _from_on_same_page(self, frame):
        """The From box's value while the page has NOT reloaded since Search; None otherwise."""
        try:
            return frame.evaluate("""() => {
                if (!document.documentElement.hasAttribute('data-bot-stale')) return null;
                const el = document.getElementById('frmDt') || document.querySelector("input[name='frmDt']");
                return el ? el.value : null; }""")
        except Exception:
            return None

    def scrape(self):
        """The results grid of the current page as a DataFrame, or None."""
        from io import StringIO
        import pandas as pd
        from engines.wht_outputs import flatten_columns
        best, best_score = None, -1
        empty_note = False
        for fr in self.frames():
            try:
                html = fr.evaluate(CLEAN_HTML_JS)       # without stale (old) results / the form
            except Exception:
                try:
                    html = fr.content()
                except Exception:
                    continue
            if re.search(NO_RECORD_RE, html, re.I):
                empty_note = True
            try:
                hits = pd.read_html(StringIO(html), match=RESULT_MATCH)
            except Exception:
                hits = []
            for tbl in hits:
                if tbl.shape[0] >= 1 and tbl.shape[1] >= 5:
                    score = 1000 + tbl.size
                    if score > best_score:
                        best, best_score = tbl, score
            if best is None:
                try:
                    tables = pd.read_html(StringIO(html))
                except Exception:
                    tables = []
                for tbl in tables:
                    if tbl.shape[0] < 1 or tbl.shape[1] < 5:
                        continue
                    flat = flatten_columns(tbl.copy())
                    head = " ".join(str(c) for c in flat.columns).lower()
                    if "search" in head or "reset" in head:
                        continue
                    n = sum(1 for k in KEYWORDS if k in head)
                    if n < 2:
                        continue
                    score = n * 1000 + tbl.size
                    if score > best_score:
                        best, best_score = tbl, score
        if best is None:
            self.log("  No records for this month." if empty_note
                     else "  No results table found on the page.")
            return None
        self.log(f"  Captured {len(best)} row(s), {len(best.columns)} column(s).")
        return best

    def scrape_paged(self, max_pages=60):
        """Follow 'Next' until the list repeats or the link disappears."""
        import pandas as pd
        frames, last_sig = [], None
        for page_no in range(1, max_pages + 1):
            self.check()
            df = self.scrape()
            if df is None or not len(df):
                break
            sig = (tuple(str(x) for x in df.iloc[0].tolist()), len(df))
            if sig == last_sig:
                self.log("  Same page again - end of the list.")
                break
            last_sig = sig
            frames.append(df)
            self.clear_overlays()
            nxt = None
            for fr in self.frames():
                for sel in SEL["next_page"]:
                    try:
                        loc = fr.locator(sel).first
                        if loc.count() and loc.is_visible() and loc.is_enabled():
                            cls = loc.get_attribute("class") or ""
                            if "disabled" not in cls:
                                nxt = loc
                                break
                    except Exception:
                        continue
                if nxt is not None:
                    break
            if nxt is None:
                break
            self.log(f"  Moving to page {page_no + 1}...")
            for fr in self.frames():
                try:
                    fr.evaluate(MARK_STALE_JS)
                except Exception:
                    pass
            try:
                nxt.click(timeout=6000)
            except Exception:
                break
            end = time.monotonic() + RESULT_WAIT_MS / 1000.0
            fresh = False
            while time.monotonic() < end and not fresh:
                self.check()
                for fr in self.frames():
                    try:
                        if fr.evaluate(FRESH_JS) == "table":
                            fresh = True
                            break
                    except Exception:
                        continue
                if not fresh:
                    self.page.wait_for_timeout(250)
            if not fresh:
                raise RuntimeError(f"page {page_no + 1} of the results did not load within "
                                   f"{RESULT_WAIT_MS // 1000} s")
        if not frames:
            return None
        return pd.concat(frames, ignore_index=True)

    # ------------------------------------------------------------ page copies
    def _expand_frames(self):
        """printToPDF renders an iframe at its CSS height, so stretch each
        frame to its content height first."""
        saved = []
        for fr in self.frames():
            if fr == self.page.main_frame:
                continue
            try:
                h = fr.evaluate("() => Math.max(document.documentElement.scrollHeight, "
                                "document.body ? document.body.scrollHeight : 0)")
                w = fr.evaluate("() => Math.max(document.documentElement.scrollWidth, "
                                "document.body ? document.body.scrollWidth : 0)")
                el = fr.frame_element()
                saved.append((el, el.get_attribute("style") or "", el.get_attribute("id") or ""))
                el.evaluate("(e, d) => { e.style.height = d.h + 'px'; e.style.minWidth = d.w + 'px'; "
                            "e.id = '__bot_main'; }", {"h": int(h) + 80, "w": int(w)})
            except Exception:
                continue
        if saved:
            self.page.wait_for_timeout(600)
        return saved

    @staticmethod
    def _restore_frames(saved):
        for el, style, el_id in saved:
            try:
                el.evaluate("(e, d) => { e.setAttribute('style', d.s); if (d.i) e.id = d.i; "
                            "else e.removeAttribute('id'); }", {"s": style, "i": el_id})
            except Exception:
                continue

    def page_pdf(self):
        """Print the results page.  page.pdf() refuses to run in a headed
        browser, so use the DevTools protocol.  The site's own print CSS is
        overridden to hide the wrapper and force landscape."""
        saved = self._expand_frames()
        style_id = None
        try:
            style_id = self.page.evaluate("""css => {
                const s = document.createElement('style');
                s.id = '__bot_print'; s.textContent = css;
                document.head.appendChild(s); return s.id; }""", PRINT_CSS)
        except Exception:
            pass
        data = None
        try:
            if self._cdp is None:
                self._cdp = self.page.context.new_cdp_session(self.page)
            res = self._cdp.send("Page.printToPDF", {
                "landscape": True, "printBackground": True,
                "paperWidth": 11.69, "paperHeight": 8.27,
                "marginTop": 0.3, "marginBottom": 0.3, "marginLeft": 0.3, "marginRight": 0.3,
                "scale": 0.75, "preferCSSPageSize": True})
            data = base64.b64decode(res["data"])
        except Exception as exc:
            self.log(f"  (page-copy PDF unavailable: {exc})")
        finally:
            if style_id:
                try:
                    self.page.evaluate("id => { const s = document.getElementById(id); "
                                       "if (s) s.remove(); }", style_id)
                except Exception:
                    pass
            self._restore_frames(saved)
        return data

    def grid_screenshot(self):
        loc = self.find(SEL["results_table"], timeout=3000)
        if loc is not None:
            try:
                return loc.screenshot(type="png")
            except Exception:
                pass
        return None

    def page_html(self):
        for fr in self.frames():
            try:
                html = fr.content()
                if "GenTaxCreditCert" in html or "<table" in html.lower():
                    return html
            except Exception:
                continue
        try:
            return self.page.content()
        except Exception:
            return ""

    # ------------------------------------------------------------ TCC numbers
    def expecting_generate_confirm(self):
        """Within this block - and only here - the portal's 'generate Tax Credit
        Certificate?' question is answered OK (engines/dialogs.py)."""
        return self._guard.expecting()

    def tcc_status(self, tick=False):
        for fr in self.frames():
            try:
                res = fr.evaluate(TCC_TICK_JS, tick)
            except Exception:
                continue
            if res and res.get("found"):
                res["frame"] = fr
                return res
        return {"found": False}

    def generate_missing_tcc(self):
        """Tick rows without a TCC number, press Generate TCC, wait for the
        summary screen and press its Back button.  Returns how many rows were
        sent for generation (0 = nothing to do).  The caller searches the month
        again to read the new TCC numbers - the portal shows them only then."""
        self.check()
        st = self.tcc_status(tick=False)
        if not st.get("found") or not st.get("missing"):
            return 0
        st = self.tcc_status(tick=True)
        if not st.get("ticked"):
            raise RuntimeError(f"{st.get('missing')} row(s) have no TCC number but none could be ticked")
        fr = st["frame"]
        self.log(f"  {st['ticked']} row(s) without a TCC number ticked - pressing Generate TCC...")
        self.dialogs.clear()
        btn = None
        for f in [fr] + [x for x in self.frames() if x is not fr]:
            for sel in SEL["generate_btn"]:
                try:
                    loc = f.locator(sel).first
                    if loc.count():
                        btn = loc
                        break
                except Exception:
                    continue
            if btn is not None:
                break
        if btn is None:
            raise RuntimeError("the Generate TCC button was not found")
        with self.expecting_generate_confirm():
            btn.evaluate("el => { setTimeout(() => el.click(), 0); }", timeout=5000)
            summary = self.wait_find(SEL["tcc_summary"], timeout=RESULT_WAIT_MS)
        for m in self.dialogs:
            self.log(f"  Portal message: {m}")
        if summary is None:
            bad = [m for m in self.dialogs if "select" in m.lower()]
            raise RuntimeError("Generate TCC: " + (bad[0] if bad else
                               f"no summary screen within {RESULT_WAIT_MS // 1000} s"))
        try:
            text = " ".join((summary.inner_text(timeout=5000) or "").split())
            self.log(f"  Summary: {text[:160]}")
        except Exception:
            pass
        back = self.wait_find(SEL["summary_back"], timeout=10000)
        if back is not None:
            back.evaluate("el => el.click()", timeout=5000)
        if self.wait_find(SEL["from_date"], timeout=30000) is None:
            self.open_wht_page(quiet=True)
        self.log(f"  Generate pressed for {st['ticked']} row(s); searching the month again "
                 "to read the numbers.")
        return st["ticked"]

    def logout(self):
        try:
            self.click_any(SEL["logout_btn"], "Logout", required=False, timeout=5000)
            self.page.wait_for_timeout(1500)
        except Exception:
            pass


# ====================================================================== loop
def run_months(bot, months, want_pdf=True, want_screenshot=True, want_html=True,
               generate_for="D", continue_on_error=True, on_result=None, generate_tcc=False):
    """Loop the requested months.  `months` is a list of MonthResult that
    are filled in place.  Calls on_result(result, index) after each."""
    from engines.wht_outputs import clean_frame, check_period
    for i, r in enumerate(months, start=1):
        started = time.time()
        bot.log(f"--- {r.title} ({i} of {len(months)}) ---")
        for attempt in (1, 2):                       # one retry, like the other modules
            r.warnings = []
            try:
                bot.check()
                if attempt == 2:
                    if ONLINE_WAIT:
                        ONLINE_WAIT()
                    bot.log(f"  Retrying {r.title} (attempt 2 of 2)...")
                    bot.open_wht_page(quiet=True)
                frame = bot.form_frame()
                bot.clear_overlays(frame)
                r.warnings += bot.pick_month(frame, r.year, r.month)
                bot.set_generate_for(frame, generate_for)
                bot.search(frame)
                raw = bot.scrape_paged()
                if generate_tcc and raw is not None and len(raw):
                    try:
                        made = bot.generate_missing_tcc()
                    except (Cancelled, SetupError):
                        raise
                    except Exception as exc:          # never lose the month's data over this
                        made = 0
                        msg = (str(exc).splitlines() or [""])[0][:150]
                        bot.log(f"  WARNING: TCC numbers not generated: {msg}")
                        r.warnings.append(f"TCC numbers not generated: {msg}")
                    if made:
                        frame = bot.form_frame()
                        bot.clear_overlays(frame)
                        r.warnings += bot.pick_month(frame, r.year, r.month)
                        bot.set_generate_for(frame, generate_for)
                        bot.search(frame)
                        raw = bot.scrape_paged()
                        left = bot.tcc_status().get("missing", 0)
                        r.warnings.append(f"Generate pressed for {made} row(s)")
                        if left:
                            r.warnings.append(f"{left} row(s) still without a TCC number")
                if raw is not None and len(raw):
                    df = clean_frame(raw)
                    bad = check_period(df, r.portal_period)
                    if bad is not None and len(bad):
                        msg = (f"{len(bad)} row(s) show a different period than "
                               f"{r.portal_period} - check the page copy")
                        bot.log(f"  WARNING: {msg}.")
                        r.warnings.append(msg)
                    r.df = df
                    r.status = "ok"
                    r.message = "; ".join(r.warnings)
                    if want_pdf:
                        r.pdf_bytes = bot.page_pdf()
                    if want_screenshot:
                        r.screenshot = bot.grid_screenshot()
                    if want_html:
                        r.html = bot.page_html()
                else:
                    r.status = "empty"
                    how = (" (the page reloaded with no results table - see the page copy)"
                           if getattr(bot, "no_table_after_reload", False) else "")
                    r.message = "; ".join(r.warnings) or "Portal returned no records" + how
                    if want_html:
                        r.html = bot.page_html()
                    bot.log(f"  Nothing recorded for {r.title}.")
                break
            except Cancelled:
                r.status, r.message = "skipped", "Stopped by the user"
                r.seconds = time.time() - started
                if on_result:
                    on_result(r, i)
                raise
            except SetupError as exc:
                bot.log(f"  STOPPED this type: {exc}")
                for rr in months[i - 1:]:
                    rr.status, rr.message = "error", str(exc)[:200]
                if on_result:
                    on_result(r, i)
                return
            except Exception as exc:
                r.status = "error"
                r.message = (str(exc).splitlines() or [""])[0][:200]
                bot.log(f"  ERROR on {r.title} (attempt {attempt} of 2): {r.message}")
                if attempt == 2:
                    shot = bot.snapshot(f"error_{r.label}")
                    if shot:
                        bot.log(f"  Screenshot saved: {shot}")
                    if not continue_on_error:
                        r.seconds = time.time() - started
                        if on_result:
                            on_result(r, i)
                        raise
        r.seconds = time.time() - started
        if on_result:
            on_result(r, i)
        if i < len(months):
            try:
                bot.open_wht_page(quiet=True)
            except Exception as exc:
                bot.log(f"  Could not reopen the form: {exc}")
                if not continue_on_error:
                    raise
