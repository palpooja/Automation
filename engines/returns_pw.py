"""
engines/returns_pw.py - Return history -> PDF / .mhtml on Playwright.

PORT of run_bot() in the Returns Downloader v4.3 (Selenium; now legacy/).  Same page flow,
same selectors (translated from XPath), same file naming, same
Returns_Status.xlsx (write_status_sheet is reused as-is: it is pure
openpyxl).  Differences:

  * runs on a Playwright `page` that is ALREADY logged in (shared session)
  * PDF: instead of clicking the portal's Print button and catching Chrome's
    silent "Save as PDF" (needed --kiosk-printing and a temp folder), the
    return tab is printed through DevTools Page.printToPDF - the same call
    the Ledger engine already uses with the window open.  'Print with
    Schedules' is still clicked first when the page has it, with
    window.print() disabled, so the schedule view is what gets printed.
  * .mhtml: DevTools Page.captureSnapshot, as before.
  * Stop is honoured at every step through check().
  * "already saved" rule: a return whose file already exists is skipped
    unless its status cell changed since the last run (speed item 4).
"""
import datetime
import os
import re
import time

from engines import returns_common as old

ALL_CATEGORIES = old.ALL_CATEGORIES
DEFAULT_TICKED = old.DEFAULT_TICKED
BRANCH_CODES = old.BRANCH_CODES
HTML_CATEGORIES = old.HTML_CATEGORIES
STATUS_XLSX = old.STATUS_XLSX
safe_name, unique_path, write_status_sheet, _clean = (old.safe_name, old.unique_path,
                                                      old.write_status_sheet, old._clean)

ROWS_WITH_VIEW = "#returnHistory tbody tr:not([data-bot-stale]):has(td:nth-child(11) a)"
NEXT_SEL = ("a:text-is('Next'), a:has-text('Next'), input[value='Next'], "
            "button:has-text('Next')")
SCHED_SEL = ("#printWthSchid, input[value*='print with sched' i], "
             "button:has-text('Print with Sched'), a:has-text('Print with Sched')")
PRINT_SEL = "#printid"
# After 'Print' is seen, keep looking for 'Print with Schedules' this long.  The
# old loop looked for 1.5 s + 1 s after it FOUND Print, and found Print up to 1 s
# late (it spent 1 s looking for Schedules first), so it watched 2.5-3.5 s after
# Print appeared.  3.5 s here is never shorter, so no return is printed without
# schedules that the old code would have printed with them.
SCHED_GRACE_S = 3.5

UNBOX_JS = """() => {
    var css = '.div1, [style*="overflow-X: scroll"], [style*="overflow-x: scroll"],'
            + '[style*="overflow-x:scroll"], [style*="overflow-X:scroll"]'
            + '{overflow:visible !important; width:auto !important;'
            + ' margin-left:0 !important; margin-right:0 !important;}'
            + '.div1 table, [style*="overflow-x"] table, [style*="overflow-X"] table'
            + '{width:auto !important; min-width:100% !important;}'
            + 'table.tab3 th, table.tab3 td {white-space:nowrap;}';
    var s = document.createElement('style');
    s.setAttribute('data-ura-unbox', '1');
    s.appendChild(document.createTextNode(css));
    (document.head || document.documentElement).appendChild(s);
    document.querySelectorAll('*').forEach(function (el) {
        var st = (el.getAttribute('style') || '').toLowerCase();
        if (st.indexOf('overflow') >= 0 && st.indexOf('scroll') >= 0) {
            el.style.setProperty('overflow', 'visible', 'important');
            el.style.setProperty('width', 'auto', 'important');
        }
    });
    return true;
}"""


MARK_STALE_JS = """() => {
    document.querySelectorAll('#returnHistory, #returnHistory *').forEach(
        e => e.setAttribute('data-bot-stale', '1'));
    document.querySelectorAll('body *').forEach(e => {
        if (/total\\s*records?\\s*:|current\\s*page\\s*:/i.test(e.textContent || '')
            && e.children.length === 0) e.setAttribute('data-bot-stale', '1');
    });
}"""
FRESH_JS = """() => {
    const area = Array.from(document.querySelectorAll('#returnHistory'));
    return area.some(a => !a.hasAttribute('data-bot-stale')
        || Array.from(a.querySelectorAll('tr')).some(r => !r.hasAttribute('data-bot-stale')));
}"""


def wait_fresh_results(frame, timeout_s, check):
    """True once results that were NOT on screen before Search are present."""
    end = time.monotonic() + timeout_s
    while time.monotonic() < end:
        check()
        try:
            if frame.evaluate(FRESH_JS):
                return True
        except Exception:
            pass
        frame.page.wait_for_timeout(250)
    return False


def fresh_text(frame):
    """Visible text of the results area excluding anything marked stale."""
    try:
        return frame.evaluate("""() => {
            const c = document.body.cloneNode(true);
            c.querySelectorAll('[data-bot-stale]').forEach(e => e.remove());
            return c.innerText || c.textContent || ''; }""")
    except Exception:
        return ""


def first_row_text(frame):
    try:
        loc = frame.locator(ROWS_WITH_VIEW).first
        return loc.inner_text(timeout=2000) if loc.count() else ""
    except Exception:
        return ""


# ---------------------------------------------------------------- file names
def stable_label(cells, fallback):
    """File name from FIXED facts: return number + period + filing date.
    The old name began with the row number, which shifts every time a new
    return is filed, so the same return was saved again under a new name."""
    norm = [" ".join(str(c).split()) for c in cells]
    retno = next((c for c in norm if re.fullmatch(r"[A-Z0-9]{10,}", c)
                  and re.search(r"[A-Z]", c) and re.search(r"\d", c)), "")
    if not retno:
        return fallback, ""
    dates = [c for c in norm if re.fullmatch(r"\d{2}/\d{2}/\d{4}", c)]
    parts = [retno]
    if len(dates) >= 2:
        parts.append(f"{dates[0].replace('/', '-')} to {dates[1].replace('/', '-')}")
    if len(dates) >= 3:
        parts.append(f"filed {dates[2].replace('/', '-')}")
    return safe_name(" ".join(parts)), retno


MIN_PDF_BYTES = 1000      # a smaller PDF is an empty print, never a saved return


def existing_downloads(folder):
    """Names (without extension) of the returns really on disk in `folder`.
    A PDF of MIN_PDF_BYTES or less, an empty web page and a left-over temporary
    file do not count - they are retried (finding H8c)."""
    out = set()
    try:
        names = os.listdir(folder)
    except OSError:
        return out
    for name in names:
        stem, ext = os.path.splitext(name)
        ext = ext.lower()
        try:
            size = os.path.getsize(os.path.join(folder, name))
        except OSError:
            continue
        if ext == ".pdf" and size <= MIN_PDF_BYTES:
            continue
        if ext in (".mhtml", ".part") and (size == 0 or ext == ".part"):
            continue
        out.add(stem)
    return out


def no_records_screenshot(page, frame, folder, date_from, date_to):
    """10.13 (user, 8 Oct 2026): evidence that the search was made - the search screen with Form For,
    Return For, the dates and the portal's 'No records' message - as
    '<type folder>\\No Records Found <2025-26>.png'.  Returns the path, or '' (never stops the run)."""
    path = os.path.join(folder, f"No Records Found {old.fy_label(date_from, date_to)}.png")
    for shoot in (lambda: frame.locator("body").first.screenshot(path=path, timeout=10000),
                  lambda: page.screenshot(path=path, full_page=True, timeout=10000)):
        try:
            shoot()
            if os.path.getsize(path) > 0:
                return path
        except Exception:
            continue
    return ""


def choose_name(folder, existing, mapping, cells, heads, return_type, period_from, ext, old_label):
    """(name without extension, return number, saved_already).  Friendly name when the row gives a period
    (10.13), else the 10.1 name.  A return saved under its 10.1-10.12 name (return number + dates) is
    renamed to the friendly name, never downloaded again."""
    facts = old.return_facts(cells, heads)
    label = old.friendly_label(cells, heads, return_type, period_from)
    if not label:
        label, retno = stable_label(cells, old_label)
        return label, retno, None
    retno = facts["retno"]
    label = old.unique_label(label, retno, mapping, facts["applied"])
    if retno and mapping.get(label) == retno and label in existing:
        return label, retno, True
    if label in existing and not mapping.get(label):
        mapping[label] = retno or label                     # on disk, e.g. its record was lost
        return label, retno, True
    if retno:
        for stem in sorted(existing):
            if retno in stem and stem != label:              # saved by an earlier version
                for e in (ext, ".pdf", ".mhtml"):
                    src = os.path.join(folder, stem + e)
                    dst = os.path.join(folder, label + e)
                    if os.path.exists(src) and not os.path.exists(dst):
                        try:
                            os.replace(src, dst)
                        except OSError:
                            return stem, retno, True          # cannot rename (open?) - still saved
                        existing.discard(stem)
                        existing.add(label)
                        mapping[label] = retno
                        return label, retno, True
    return label, retno, False


def already_saved(existing, label, retno):
    """True if this return is on disk under the new OR the old naming."""
    if label in existing:
        return True
    return bool(retno) and any(retno in name for name in existing)


# ---------------------------------------------------------------- helpers
DEBUG = bool(os.environ.get("URA_DEBUG"))
ONLINE_WAIT = None     # suite: waits while the internet is down


def _frames(page):
    try:
        return list(page.frames)
    except Exception:
        return [page]


def _find(page, sel, timeout_s=10, check=None, visible=True):
    """First matching element in any frame, polling every 0.25 s."""
    end = time.monotonic() + timeout_s
    while True:
        if check:
            check()
        for fr in _frames(page):
            try:
                loc = fr.locator(sel).first
                if loc.count() and (not visible or loc.is_visible()):
                    return fr, loc
            except Exception:
                continue
        if time.monotonic() >= end:
            return None, None
        page.wait_for_timeout(250)


PRINT_CATCHER = "window.print = function () { window.__botPrint = true; };"


def install_print_catcher(context):
    """window.print() would open the print dialog in a visible browser, so it is replaced by a
    function that only records the call.  That record matters: the portal calls print() only when
    its view - with the schedules - is complete (the old Selenium program 4.3 made its PDF from
    exactly that call)."""
    context.add_init_script(PRINT_CATCHER)


def _print_requested(tab):
    for fr in _frames(tab):
        try:
            if fr.evaluate("() => window.__botPrint === true"):
                return True
        except Exception:
            continue
    return False


# Return types for which URA no longer gives schedules (confirmed by the user, 4 Oct 2026).
# 'Print with Schedules' is still pressed for them, so schedules are captured if URA brings them back.
NO_SCHEDULES_FROM_URA = {"Monthly VAT Return"}
READY = "Print with Schedules - printed when the portal asked to print"
QUIET_S = 5.5          # = the old fixed wait (4 s for a new tab + 1.5 s): never slower than before
NEXT_PAGE_WAIT_S = 60          # normal mode: how long a page turn may take.  Live page turns took 6-20 s
                               # (5 Oct 2026); the old fixed 40 x 0.3 s (about 12-13 s) lost client B's
                               # second page of VAT returns.
NEXT_PAGE_WAIT_PATIENT_S = 180 # 'wait until it downloads'
SCHED_CAP_PATIENT_S = 300   # upper limit for the schedules; normally the page limit (60 s x speed) is lower


class _Activity:
    """Requests in flight on a tab - the portal still fetching its schedules."""
    def __init__(self, tab):
        self.inflight, self.last = 0, time.monotonic()
        self._tab = tab
        tab.on("request", self._start)
        tab.on("requestfinished", self._end)
        tab.on("requestfailed", self._end)

    def _start(self, _):
        self.inflight += 1
        self.last = time.monotonic()

    def _end(self, _):
        self.inflight = max(0, self.inflight - 1)
        self.last = time.monotonic()

    def close(self):
        for ev, fn in (("request", self._start), ("requestfinished", self._end), ("requestfailed", self._end)):
            try:
                self._tab.remove_listener(ev, fn)
            except Exception:
                pass


def press_print_with_schedules(context, tab, sched, limit_s, check=lambda: None,
                               log=lambda *a: None, tag="", expects_schedules=True):
    """Press 'Print with Schedules', then make the PDF:
      * as soon as the portal asks to print (its view is complete) - on this tab or a new one;
      * otherwise once the page has been quiet - nothing loading, nothing changing - for QUIET_S
        (the old fixed wait, so never slower and never less complete than before);
      * at the latest after limit_s, flagged for a person.
    Returns (page to print, detail for the summary).  Details starting 'check' need a person."""
    for fr in _frames(tab):
        try:
            fr.evaluate("() => { window.__botPrint = false; }")
        except Exception:
            continue
    before = list(context.pages)
    activity = _Activity(tab)
    try:
        _js_click(sched)
        t0 = time.monotonic()
        size, quiet_since, next_note = _text_size(tab), time.monotonic(), 30
        while time.monotonic() - t0 < limit_s:
            check()
            if _print_requested(tab):
                tab.wait_for_timeout(300)            # let the print layout settle
                return tab, READY
            for p_ in [p for p in context.pages if p not in before]:
                try:
                    p_.wait_for_load_state("load", timeout=10000)
                except Exception:
                    pass
                end = time.monotonic() + 1.5
                while time.monotonic() < end and not _print_requested(p_):
                    p_.wait_for_timeout(250)
                return p_, READY + " (new tab)"
            now_size = _text_size(tab)
            if now_size != size or activity.inflight:
                size, quiet_since = now_size, time.monotonic()
            if time.monotonic() - max(quiet_since, activity.last) >= QUIET_S:
                if expects_schedules:
                    return tab, ("check: 'Print with Schedules' pressed, but the portal neither asked to "
                                 "print nor changed the page - printed what was on screen")
                return tab, ("Print with Schedules pressed; URA no longer adds schedules to this "
                             "return type - printed as shown")
            waited = time.monotonic() - t0
            if waited >= next_note:
                log(f"   {tag} still waiting for the portal to prepare the schedules ({int(waited)} s)...")
                next_note += 30
            tab.wait_for_timeout(250)
        return tab, (f"check: the portal was still working on the schedules after {int(limit_s)} s - "
                     "printed what was on screen")
    finally:
        activity.close()


def _text_size(tab):
    """Amount of visible text in all frames - grows when the schedules are added."""
    total = 0
    for fr in _frames(tab):
        try:
            total += fr.evaluate("() => (document.body && document.body.innerText || '').length")
        except Exception:
            continue
    return total


def wait_for_print_buttons(tab, page_limit, check, log, tag="", tick=None):
    """Wait until the return shows a print button.  Returns (loaded, frame, schedules_button).
    Both buttons are looked for in one pass every 0.25 s; after 'Print' is seen,
    'Print with Schedules' is looked for up to SCHED_GRACE_S and taken as soon as it appears."""
    end = time.monotonic() + page_limit
    t0, next_note = time.monotonic(), 30
    while time.monotonic() < end:
        check()
        waited = time.monotonic() - t0
        if waited >= next_note:                 # heartbeat: never look frozen
            log(f"   {tag} still waiting for the return to finish loading "
                f"({int(waited)} s of {page_limit} s)...")
            while next_note <= waited:          # one line even after a slow check (7 Oct: printed twice)
                next_note += 30
        if tick is not None and tick():
            return False, None, None
        tfr, sched = _find(tab, SCHED_SEL, 0)
        if sched is not None:
            return True, tfr, sched
        _, plain = _find(tab, PRINT_SEL, 0)
        if plain is not None:
            tfr, sched = _find(tab, SCHED_SEL, SCHED_GRACE_S)
            return True, tfr, sched
        tab.wait_for_timeout(250)
    return False, None, None


def wait_table_changed(page, frame, before_sig, limit_s, check, log, page_wanted):
    """True once the first row differs from before_sig (the next page is showing); False after
    limit_s.  A progress line every 30 s, never repeated when one check is slow."""
    t0 = time.monotonic()
    next_note = 30
    while time.monotonic() - t0 < limit_s:
        check()
        page.wait_for_timeout(300)
        sig = first_row_text(frame)
        if sig and sig != before_sig:
            return True
        waited = time.monotonic() - t0
        if waited >= next_note:
            log(f"   still waiting for page {page_wanted} ({int(waited)} s of {limit_s} s)...")
            while next_note <= waited:
                next_note += 30
    return False


def _js_click(loc):
    try:
        loc.scroll_into_view_if_needed(timeout=2000)
    except Exception:
        pass
    loc.evaluate("el => el.click()")


def _skip_popups(page, log):
    for fr in _frames(page):
        for sel in ("button:text-is('Skip')", "input[value='Skip']", "#parishMsgA button",
                    "#NINBRNMsg button", "#myModal button"):
            try:
                loc = fr.locator(sel).first
                if loc.count() and loc.is_visible():
                    _js_click(loc)
                    log("   (skipped a popup message)")
                    page.wait_for_timeout(400)
            except Exception:
                continue


def fill_dates(page, frame, date_from, date_to, log):
    """Fill the two date boxes (#fromDate / #toDate, else the first two
    text boxes whose id/name contains 'date', else the first two)."""
    boxes = []
    if frame.locator("#fromDate").count() and frame.locator("#toDate").count():
        boxes = [frame.locator("#fromDate").first, frame.locator("#toDate").first]
    if len(boxes) < 2:
        cand = frame.locator("input[type='text'][id*='date' i], input[type='text'][name*='date' i]")
        boxes = [cand.nth(i) for i in range(min(2, cand.count()))]
    if len(boxes) < 2:
        cand = frame.locator("input[type='text']:visible")
        boxes = [cand.nth(i) for i in range(min(2, cand.count()))]
    if len(boxes) < 2:
        log("   Could not find the date boxes in any frame.")
        return False
    for box, value in ((boxes[0], date_from), (boxes[1], date_to)):
        box.evaluate("(el, v) => { el.value = v;"
                     "el.dispatchEvent(new Event('input',{bubbles:true}));"
                     "el.dispatchEvent(new Event('change',{bubbles:true})); }", value)
    return True


def select_both_branches(frame, log):
    dd = frame.locator("#branchDrpDwn").first
    try:
        dd.wait_for(state="visible", timeout=4000)
    except Exception:
        log("   Headquarter/Both box did not appear - continuing with the portal's "
            "default (Headquarter).")
        return False
    before = dd.evaluate("el => el.options[el.selectedIndex]?.text || ''").strip() or "(blank)"
    if before.lower() == "both":
        log("   Headquarter/Both box is already 'Both'.")
        return True
    try:
        dd.select_option("B")
    except Exception:
        try:
            dd.select_option(label="Both")
        except Exception:
            pass
    after = dd.evaluate("el => el.options[el.selectedIndex]?.text || ''").strip()
    if after.lower() == "both":
        log(f"   Headquarter/Both box: '{before}' -> 'Both'")
        return True
    log(f"   Could not switch the Headquarter/Both box to 'Both' (it still says "
        f"'{after}') - continuing with '{after}'.")
    return False


def read_history_headings(frame):
    heads = []
    for sel in ("#returnHistory tr:has(th) >> nth=0 >> th", "#returnHistory thead tr >> nth=0 >> td"):
        try:
            cells = frame.locator(sel)
            n = cells.count()
        except Exception:
            n = 0
        if n:
            heads = [_clean(cells.nth(i).inner_text()) or _clean(cells.nth(i).text_content())
                     for i in range(n)]
            break
    heads = heads[:10] + [""] * max(0, 10 - len(heads))
    out, seen = [], set()
    for i, h in enumerate(heads):
        h = h or f"Column {i + 1}"
        base, n = h, 2
        while h in seen or h in (old.BOT_TYPE, old.BOT_DL, old.BOT_CHK, old.BOT_FILE):
            h = f"{base} ({n})"
            n += 1
        seen.add(h)
        out.append(h)
    return out


def read_row_cells(row):
    tds = row.locator("> td") if row.locator("> td").count() else row.locator("td")
    n = tds.count()
    out = []
    for i in range(min(10, n)):
        t = _clean(tds.nth(i).inner_text())
        if not t:
            try:
                t = _clean(tds.nth(i).text_content())
            except Exception:
                t = ""
        out.append(t)
    return out + [""] * max(0, 10 - len(out))


# printable width of an A4 page with 0.3-inch margins, in CSS pixels (96 per inch)
PORTRAIT_PX = int((8.27 - 0.6) * 96)
LANDSCAPE_PX = int((11.69 - 0.6) * 96)
# buttons the portal shows on screen; they carry no data and are hidden in print only
HIDE_CONTROLS_CSS = ("@media print { #printid, #printWthSchid, #backbtn, "
                     "input[type=button][value*='print' i], input[type=submit][value*='print' i], "
                     "input[type=button][value*='back' i], input[type=button][value*='close' i], "
                     "button[onclick*='print' i] { display: none !important; } }")
_WIDTH_JS = "() => Math.max(document.documentElement.scrollWidth, document.body ? document.body.scrollWidth : 0)"


def _page_layout(tab):
    """(landscape, scale) so that the widest content fits the page (user, 4 Oct 2026: a wide
    schedule lost its right-hand columns).  Measured with print styles at the page's printable
    width: a page that fits gets (False, 1) - exactly as before."""
    for fr in _frames(tab):
        try:
            fr.evaluate("""css => { if (document.querySelector('style[data-bot=print]')) return;
                const s = document.createElement('style'); s.setAttribute('data-bot', 'print');
                s.textContent = css; (document.head || document.documentElement).appendChild(s); }""",
                        HIDE_CONTROLS_CSS)
        except Exception:
            continue
    viewport = tab.viewport_size or {"width": 1400, "height": 900}
    try:
        tab.emulate_media(media="print")
        tab.set_viewport_size({"width": PORTRAIT_PX, "height": viewport["height"]})
        width = tab.evaluate(_WIDTH_JS)
        if width <= PORTRAIT_PX + 2:
            return False, 1.0
        tab.set_viewport_size({"width": LANDSCAPE_PX, "height": viewport["height"]})
        width = tab.evaluate(_WIDTH_JS)
        return True, (1.0 if width <= LANDSCAPE_PX + 2 else max(0.1, round(LANDSCAPE_PX / width, 3)))
    except Exception:
        return False, 1.0
    finally:
        try:
            tab.set_viewport_size(viewport)
            tab.emulate_media(media=None)
        except Exception:
            pass


def _print_to_pdf(target_page, dest):
    landscape, scale = _page_layout(target_page)
    session = target_page.context.new_cdp_session(target_page)
    try:
        res = session.send("Page.printToPDF", {"printBackground": True, "paperWidth": 8.27,
                                               "paperHeight": 11.69, "marginTop": 0.3,
                                               "marginBottom": 0.3, "marginLeft": 0.3,
                                               "marginRight": 0.3, "preferCSSPageSize": True,
                                               "landscape": landscape, "scale": scale})
        import base64
        data = base64.b64decode(res["data"])
    finally:
        try:
            session.detach()
        except Exception:
            pass
    _write_whole(dest, data)
    return dest


def _write_whole(dest, data):
    """Write to <dest>.part, then rename: a stopped run never leaves half a file
    under the final name (which a later run would take for a saved return)."""
    tmp = dest + ".part"
    mode = "wb" if isinstance(data, (bytes, bytearray)) else "w"
    kw = {} if mode == "wb" else {"encoding": "utf-8", "newline": ""}
    with open(tmp, mode, **kw) as fh:
        fh.write(data)
    os.replace(tmp, dest)


def _capture_mhtml(target_page, dest):
    for fr in _frames(target_page):
        try:
            fr.evaluate(UNBOX_JS)
        except Exception:
            pass
    target_page.wait_for_timeout(500)
    session = target_page.context.new_cdp_session(target_page)
    try:
        res = session.send("Page.captureSnapshot", {"format": "mhtml"})
    finally:
        try:
            session.detach()
        except Exception:
            pass
    data = res.get("data", "")
    if not data:
        return False
    _write_whole(dest, data)
    return True


# ---------------------------------------------------------------- main
def run_on_page(page, login_id, date_from, date_to, wanted_categories, log, check,
                base_folder, patient=False, result=None, skip_existing=True, first_try_s=None,
                person=None):
    """first_try_s: a return not ready within this many seconds is SET ASIDE (result['deferred']) instead of
    holding up the run; the caller tries it again later without first_try_s.  person: on that later try, an
    object with attention(message) / poll_answer() / hide_question() - 'Attention needed' and Skip / Keep
    waiting after SKIP_OFFER_AFTER_SECONDS (the wait goes on while the question shows)."""
    """Download every return in the period on an already-logged-in page.
    `check` raises when Stop is pressed.  Fills `result` like run_bot did:
    summary [(category, files)], ok, folder."""
    summary = []
    if result is None:
        result = {}
    result["summary"], result["ok"], items = summary, False, []
    result["items"] = items          # [{item, status, detail, file}]
    deferred = result.setdefault("deferred", [])        # [{type, label, retno}] - set aside (10.14)
    last = {"why": ""}                                  # why the last download_one gave up
    tin_folder = os.path.join(base_folder, safe_name(login_id))
    result["folder"] = tin_folder
    os.makedirs(tin_folder, exist_ok=True)
    status_rows, checked_on = [], datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    tab_limit = old.PATIENT_LIMIT_SECONDS if patient else old.TAB_WAIT_SECONDS
    page_limit = old.PATIENT_LIMIT_SECONDS if patient else old.PAGE_WAIT_SECONDS
    retries = 3 if patient else 1
    log("Mode: " + ("WAIT UNTIL IT DOWNLOADS" if patient else "normal waits") + ".")

    install_print_catcher(page.context)

    # ---- open 'View return history' ----
    link_sel = "a:has-text('return history')"
    fr, link = None, None
    for _ in range(40):
        check()
        _skip_popups(page, log)
        fr, link = _find(page, link_sel, 1, check, visible=False)
        if link is not None:
            break
    if link is None:
        raise RuntimeError("Could not find 'View return history' in the portal menu.")
    log("Opening 'View return history'...")
    href = (link.get_attribute("href") or "")
    _js_click(link)
    fr, btn = _find(page, "#search", 25, check)
    if btn is None and href and "javascript" not in href.lower():
        log("   Click did not open the page - trying its direct address...")
        page.goto(href)
        fr, btn = _find(page, "#search", 20, check)
    if btn is None:
        raise RuntimeError("The return-history search page did not load.")
    log("Return history page is open.")

    dd = fr.locator("#rtnType").first
    options = dd.evaluate("el => Array.from(el.options).map(o => [o.value, o.text.trim()])")
    categories = [(v, t) for v, t in options if v and "select" not in t.lower()]
    log(f"Portal offers {len(categories)} return types: " + ", ".join(t for _, t in categories))
    if wanted_categories is not None:
        wanted = {w.lower() for w in wanted_categories}
        new_types = [t for _, t in categories if t.lower() not in {c.lower() for c in ALL_CATEGORIES}]
        if new_types:
            log("NOTE: The portal now has new return type(s) not in the tick list: "
                + ", ".join(new_types))
        categories = [(v, t) for v, t in categories if t.lower() in wanted]
        if not categories:
            raise RuntimeError("None of your ticked categories matched the portal's list.")
        log("Downloading only: " + ", ".join(t for _, t in categories))

    # ---- one return ----
    cat_name = ""

    def dbg(msg):
        if DEBUG:
            log(f"   [debug {time.strftime('%H:%M:%S')}] {msg}")

    def download_one(view_link, folder, row_label, tag, save_html):
        started = time.time()
        new_tab = None
        last["why"] = ""
        t_lim = min(tab_limit, first_try_s) if first_try_s else tab_limit
        offer = {"shown": False, "done": False}

        def tick():
            """Second try only: after SKIP_OFFER_AFTER_SECONDS, 'Attention needed' + Skip / Keep waiting.
            True when the person pressed Skip."""
            if person is None or offer["done"]:
                return False
            if not offer["shown"] and time.time() - started >= old.SKIP_OFFER_AFTER_SECONDS:
                offer["shown"] = True
                person.attention(f"{cat_name}: {row_label} is still not open after "
                                 f"{int(time.time() - started)} s. Skip it (it is tried again on the next run) "
                                 "or keep waiting?")
            if offer["shown"]:
                ans = person.poll_answer()
                if ans in ("s", "q"):
                    offer["done"] = True
                    person.hide_question()
                    last["why"] = "skipped"
                    log(f"   {tag} skipped by you after {int(time.time() - started)} s.")
                    return True
                if ans is not None:                       # Keep waiting
                    offer["done"] = True
                    person.hide_question()
                    log(f"   {tag} you chose to keep waiting.")
            return False
        dbg(f"{tag} clicking View")
        # Watch for the new tab ourselves (instead of one long expect_page) so a
        # slow portal shows a heartbeat - live run: some returns took ~60 s to open.
        before = set(page.context.pages)
        try:
            _js_click(view_link)
        except Exception as err:
            log(f"   {tag} could not press View ({type(err).__name__})")
            return False
        next_note = 30
        while new_tab is None:
            check()
            fresh = [p_ for p_ in page.context.pages if p_ not in before]
            if fresh:
                new_tab = fresh[0]
                break
            waited = time.time() - started
            if waited > t_lim:
                log(f"   {tag} the return did not open in a new tab within {int(waited)} s")
                last["why"] = "slow"
                return False
            if tick():
                return False
            if waited >= next_note:
                log(f"   {tag} waiting for the portal to open the return ({int(waited)} s)...")
                next_note += 30
            page.wait_for_timeout(250)
        ok = False
        dbg(f"{tag} tab opened")
        try:
            p_lim = page_limit
            if first_try_s:                              # one budget from View to the Print button
                p_lim = max(5, min(page_limit, int(first_try_s - (time.time() - started))))
            new_tab.wait_for_load_state("domcontentloaded", timeout=p_lim * 1000)
            dbg(f"{tag} tab DOM loaded")
            # wait for a print button = the return has finished loading
            loaded, tfr, sched = wait_for_print_buttons(new_tab, p_lim, check, log, tag, tick=tick)
            dbg(f"{tag} print-button wait over, loaded={loaded}")
            if not loaded:
                if last["why"] != "skipped":
                    last["why"] = "slow"
                    log(f"   {tag} the return did not finish loading within {p_lim} s "
                        "(no Print button) - not saved")
                return False
            if save_html:
                target = unique_path(folder, row_label, ".mhtml")
                ok = _capture_mhtml(new_tab, target)
                log(f"   {tag} web page saved -> {os.path.basename(target)}" if ok
                    else f"   {tag} web-page copy came back empty - not saved")
                if ok:
                    items.append({"item": f"{cat_name}: {row_label}", "status": "saved",
                                  "detail": "web page (.mhtml)", "file": target})
            else:
                print_page = new_tab
                schedules = "plain print"
                if sched is not None:
                    log(f"   {tag} using 'Print with Schedules'")
                    # page_limit is already x speed setting: 60 s normally, 15 min when patient -> max 5 min
                    cap = min(page_limit, SCHED_CAP_PATIENT_S)
                    print_page, schedules = press_print_with_schedules(
                        page.context, new_tab, sched, cap, check, log, tag,
                        expects_schedules=cat_name not in NO_SCHEDULES_FROM_URA)
                    if schedules.startswith("check"):
                        log(f"   {tag} {schedules[7:]}")
                else:
                    log(f"   {tag} no 'Print with Schedules' on this return - using 'Print'")
                for f_ in _frames(print_page):
                    try:
                        f_.evaluate(UNBOX_JS)
                    except Exception:
                        pass
                target = unique_path(folder, row_label, ".pdf")
                _print_to_pdf(print_page, target)
                ok = os.path.getsize(target) > MIN_PDF_BYTES
                log(f"   {tag} saved -> {os.path.basename(target)}" if ok
                    else f"   {tag} PDF came back empty - deleted, will be retried")
                if ok:
                    items.append({"item": f"{cat_name}: {row_label}", "status": "saved",
                                  "detail": schedules, "file": target})
                    if schedules.startswith("check"):
                        items.append({"item": f"{cat_name}: {row_label}", "status": "check",
                                      "detail": schedules[7:] + " - open the PDF and check the schedules",
                                      "file": ""})
                else:
                    try:
                        os.remove(target)
                    except OSError:
                        pass
        except Exception as err:
            if "Timeout" in type(err).__name__ and not last["why"]:
                last["why"] = "slow"
            log(f"   {tag} problem ({type(err).__name__}: {str(err)[:120]})")
        finally:
            dbg(f"{tag} closing extra tabs")
            for p_ in list(page.context.pages):
                if p_ is not page:
                    try:
                        p_.close()
                    except Exception:
                        pass
            dbg(f"{tag} tabs closed")
        return ok

    stopped, why = False, None
    try:
        # ---- categories ----
        for code, name in categories:
            check()
            cat_name = name
            folder = os.path.join(tin_folder, safe_name(name))
            os.makedirs(folder, exist_ok=True)
            save_html = name.strip().lower() in {c.lower() for c in HTML_CATEGORIES}
            log(f"\n=== {name} ===" + ("   (saved as web page .mhtml, no PDF)" if save_html else ""))
            fr, _ = _find(page, "#rtnType", 15, check)
            if fr is None or not fill_dates(page, fr, date_from, date_to, log):
                log("   Skipping this category (no date boxes).")
                items.append({"item": name, "status": "failed", "detail": "no date boxes on the page"})
                continue
            fr.locator("#rtnType").first.select_option(code)
            if code in BRANCH_CODES:
                select_both_branches(fr, log)
            try:
                fr.evaluate(MARK_STALE_JS)
            except Exception:
                pass
            _js_click(fr.locator("#search").first)
            area = fr if wait_fresh_results(fr, max(15, page_limit // 2), check) else None
            if area is None:
                log("   No new results appeared after Search - skipping this category.")
                items.append({"item": name, "status": "failed", "detail": "no results area appeared"})
                summary.append((name, 0))
                continue
            total_records = None
            m = re.search(r"Total\s*Records?\s*:\s*(\d+)", fresh_text(fr))
            if m:
                total_records = int(m.group(1))
                log(f"   Portal reports {total_records} record(s) in total.")
            try:
                headings = read_history_headings(fr)
            except Exception:
                headings = [f"Column {i + 1}" for i in range(10)]
            existing = existing_downloads(folder)
            saved_as = old.load_saved_as(folder)
            aside = 0

            def set_aside(rn, tag, label, retno):
                """First try only: a return that was too slow waits for the second try (10.14)."""
                nonlocal aside
                if not first_try_s or last["why"] != "slow":
                    return False
                aside += 1
                deferred.append({"type": name, "label": label, "retno": retno})
                log(f"   {tag} not ready within {first_try_s} s - set aside; it is tried again with the full "
                    "wait after this client's other steps.")
                items.append({"item": f"{name}: {label}", "status": "set aside",
                              "detail": f"not ready within {first_try_s} s - second try at the end of the client"})
                return True

            def gave_up(rn, label):
                items.append({"item": f"{name}: {label}", "status": "failed",
                              "detail": "skipped by you while waiting - tried again on the next run"})
            ext = ".mhtml" if save_html else ".pdf"
            downloaded = skipped = 0
            page_no = 1
            while True:
                check()
                rows = fr.locator(ROWS_WITH_VIEW)
                n = rows.count()
                if not n:
                    if page_no == 1:
                        shot = no_records_screenshot(page, fr, folder, date_from, date_to)
                        log("   No returns found for this period."
                            + (f" Screenshot: {os.path.basename(shot)}" if shot else ""))
                        items.append({"item": name, "status": "no records",
                                      "detail": f"{date_from} to {date_to}", "file": shot})
                    break
                log(f"   Page {page_no}: {n} return(s) on this page.")
                entries = {}
                for rn in range(n):
                    try:
                        cells = read_row_cells(rows.nth(rn))
                    except Exception:
                        cells = [""] * 10
                    entries[rn] = {"type": name, "heads": headings, "cells": cells, "downloaded": False}
                    status_rows.append(entries[rn])
                failed = []
                for rn in range(n):
                    check()
                    tag = f"[page {page_no}, row {rn + 1}]"
                    try:
                        row = fr.locator(ROWS_WITH_VIEW).nth(rn)
                        old_label = safe_name(" ".join(row.inner_text().split())[:60]) \
                            or f"p{page_no}_return_{rn + 1}"
                        row_label, retno, on_disk = choose_name(
                            folder, existing, saved_as, entries[rn]["cells"], headings, name, date_from,
                            ext, old_label)
                        entries[rn]["label"] = row_label
                        entries[rn]["retno"] = retno
                        if on_disk is None:                  # no period in the row: the 10.1 rule
                            on_disk = already_saved(existing, row_label, retno)
                        if skip_existing and on_disk:
                            log(f"   {tag} already saved - skipped ({row_label})")
                            entries[rn]["downloaded"] = True
                            entries[rn]["saved_as"] = row_label + ext
                            skipped += 1
                            items.append({"item": f"{name}: {row_label}", "status": "skipped",
                                          "detail": "already on disk"})
                            continue
                        view = row.locator("td:nth-child(11) a").first
                        if download_one(view, folder, row_label, tag, save_html):
                            downloaded += 1
                            entries[rn]["downloaded"] = True
                            entries[rn]["saved_as"] = row_label + ext
                            if retno:
                                saved_as[row_label] = retno
                                old.save_saved_as(folder, saved_as)
                        elif set_aside(rn, tag, row_label, retno):
                            pass
                        elif last["why"] == "skipped":
                            gave_up(rn, row_label)
                        else:
                            log(f"   {tag} not saved - will retry at end of page")
                            failed.append(rn)
                    except Exception as err:
                        log(f"   {tag} problem ({type(err).__name__}) - will retry at end of page")
                        failed.append(rn)
                for rr in range(1, retries + 1):
                    if not failed:
                        break
                    if ONLINE_WAIT:
                        ONLINE_WAIT()
                    log(f"   Retry round {rr} of {retries} for {len(failed)} row(s) on page {page_no}...")
                    still = []
                    for rn in failed:
                        check()
                        tag = f"[RETRY {rr} page {page_no}, row {rn + 1}]"
                        try:
                            row = fr.locator(ROWS_WITH_VIEW).nth(rn)
                            old_label = safe_name(" ".join(row.inner_text().split())[:60]) \
                                or f"p{page_no}_return_{rn + 1}"
                            row_label = entries[rn].get("label") or stable_label(entries[rn]["cells"],
                                                                                 old_label)[0]
                            if not download_one(row.locator("td:nth-child(11) a").first, folder,
                                                row_label, tag, save_html):
                                if set_aside(rn, tag, row_label, entries[rn].get("retno", "")):
                                    continue
                                if last["why"] == "skipped":
                                    gave_up(rn, row_label)
                                    continue
                                still.append(rn)
                            else:
                                downloaded += 1
                                entries[rn]["downloaded"] = True
                                entries[rn]["saved_as"] = row_label + ext
                                retno = old.return_facts(entries[rn]["cells"], headings)["retno"]
                                if retno and old.friendly_label(entries[rn]["cells"], headings, name, date_from):
                                    saved_as[row_label] = retno
                                    old.save_saved_as(folder, saved_as)
                        except Exception as err:
                            log(f"   {tag} problem again ({type(err).__name__})")
                            still.append(rn)
                    failed = still
                    if failed and rr < retries:
                        page.wait_for_timeout(5000)
                for rn in failed:
                    log(f"   *** [page {page_no}, row {rn + 1}] could not be downloaded after "
                        "all retries - skipped ***")
                    items.append({"item": f"{name}: page {page_no} row {rn + 1} "
                                          f"({' '.join(entries[rn]['cells'][:3])})",
                                  "status": "failed", "detail": "not downloaded after all retries"})
                # next page?
                txt = fresh_text(fr)
                pinfo = re.search(r"Current\s*Page\s*:\s*(\d+)\s*/\s*(\d+)", txt)
                if pinfo and int(pinfo.group(1)) >= int(pinfo.group(2)):
                    log(f"   Last page reached ({pinfo.group(1)}/{pinfo.group(2)}).")
                    break
                before_sig = first_row_text(fr)
                nxt = fr.locator(NEXT_SEL).first
                if not nxt.count():
                    log("   No 'Next' button found - treating this as the last page.")
                    break
                log(f"   Turning to page {page_no + 1}...")
                _js_click(nxt)
                turn_limit = NEXT_PAGE_WAIT_PATIENT_S if patient else NEXT_PAGE_WAIT_S
                if not wait_table_changed(page, fr, before_sig, turn_limit, check, log, page_no + 1):
                    log(f"   'Next' was clicked but the table did not change within {turn_limit} s - "
                        "treating this as the last page.")
                    break
                page_no += 1
            if saved_as:
                old.save_saved_as(folder, saved_as)
            if total_records is not None:
                got = downloaded + skipped
                if aside and got + aside >= total_records:
                    log(f"   {got} of {total_records} record(s) on disk; {aside} set aside for the second try.")
                elif got < total_records:
                    log(f"   *** WARNING: portal reported {total_records} record(s) but only {got} "
                        "file(s) are on disk. Some returns may be missing - please re-check this "
                        "category. ***")
                    items.append({"item": name, "status": "incomplete",
                                  "detail": f"portal reported {total_records}, {got} on disk"})
                else:
                    log(f"   All {total_records} record(s) on disk "
                        f"({downloaded} new, {skipped} already saved) - complete.")
            summary.append((name, downloaded))

    except BaseException as exc:
        stopped, why = True, exc
        raise
    finally:
        if stopped:
            from core.runner import Stopped, browser_gone
            if browser_gone(None, why):
                log("The browser window was closed - writing the status sheet for the returns processed so far.")
            elif isinstance(why, Stopped):
                log("Stopped - writing the status sheet for the returns processed so far.")
            else:
                log("Interrupted by an error - writing the status sheet for the returns processed so far.")
            try:
                info = write_status_sheet(tin_folder, status_rows, checked_on, log)
                if info:
                    items.append({"item": "Returns_Status.xlsx (partial, stopped)",
                                  "status": "saved", "file": info["path"]})
            except Exception:
                pass
    log("\n==================== FINISHED ====================")
    for n_, c_ in summary:
        log(f"   {n_:35s} {c_} file(s)")
    log(f"All files are in: {tin_folder}")
    try:
        info = write_status_sheet(tin_folder, status_rows, checked_on, log)
        if info:
            log(f"Status sheet: {os.path.basename(info['path'])}  ({len(status_rows)} return(s) "
                f"this run, {info['total']} listed in all - see the Summary tab)")
            items.append({"item": "Returns_Status.xlsx", "status": "saved", "file": info["path"]})
    except Exception as err:
        log(f"ERROR: could not write {STATUS_XLSX} ({type(err).__name__}: {err})")
    result["ok"] = True
    return result
