"""
engines/paye_engine.py - 'Generate Tax Credit Certificate for Pay As You Earn'.

Built from the live recording of 3 Oct 2026 4:23 PM (tests/live_recording_paye_2026-10-03.json):
  menu link   'Generate Tax Credit Certificate for Pay As You Earn' (#mailbox)
  form        #TaxCreditCertForm: #retBrnch (Return For: H = Headquarter),
              #period (value '2025-2026' = 'July 2025 to June 2026'), Search button
  results     Select All | Sr No. | Employee Tin | Employee Name | Chargeable Income |
              Tax Deducted | TCC No. | 'Get Month-Wise Details' link; row tick box #listPaye
  month-wise  the link opens a SEPARATE POP-UP WINDOW (OpenWindowWithPostWithParam -> getMonthsDtls) holding
              #divPostContent with a Close button (3 Oct recording steps 28-35 run in that window; the user's
              capture of 8 Oct 2026 shows the window.open).  Read there, then the window is closed.
  generate    'GenerateTCC' button; the portal gives no PDF - the TCC No. shows on a new search.  The page allows
              at most `count` (5) ticks per press - a 6th is refused with an alert - so it is done in batches.
  search      a real form POST: the whole PAYE page is replaced.  Only the frame holding the PAYE form counts
              as the answer (8 Oct 2026: any other frame reloading - a chat or tracking widget - was taken as
              'the portal answered without a list' and gave 'No employees listed').

For each financial year (July-June) touched by the period:
  search -> tick employees with no TCC No. -> GenerateTCC -> search again ->
  open every employee's month-wise details -> one Excel workbook per year:
  sheet 'Employees' (with TCC numbers) and sheet 'Month-wise details' (all employees).
"""
import io
import re
import time
from pathlib import Path

MENU_ITEM = "Generate Tax Credit Certificate for Pay As You Earn"
SEL = {
    "branch": ["#retBrnch", "select[name='retBrnch']"],
    "period": ["#period", "select[name='period']"],
    "search": ["#TaxCreditCertForm button:has-text('Search')", "#TaxCreditCertForm input[value='Search']",
               "button:has-text('Search')", "input[value='Search']"],
    "generate": ["input[value='GenerateTCC']", "input[value*='GenerateTCC' i]",
                 "input[value*='Generate' i]", "button:has-text('Generate')"],
    "popup": ["#divPostContent"],
    "popup_close": ["#divPostContent input[value='Close']", "#divPostContent input[type=button]",
                    "#divPostContent button:has-text('Close')"],
    "processing_close": ["#process-static-modal button.btn-secondary"],
}
WAIT_S = 60
PRESS_AGAIN_S = 8        # 10.15: Search sent nothing to the portal in this time -> press it once more
RELOAD_GRACE_S = 10      # after the PAYE page has reloaded: how long a list may take to draw (ends when it shows)
MONTHWISE_WAIT_S = 30


def financial_years(period_from, period_to):
    """'2025-05', '2026-06' -> ['2024-2025', '2025-2026'] (Uganda: July to June)."""
    y1, m1 = (int(x) for x in period_from.split("-"))
    y2, m2 = (int(x) for x in period_to.split("-"))
    out, y, m = [], y1, m1
    while (y, m) <= (y2, m2):
        start = y if m >= 7 else y - 1
        fy = f"{start}-{start + 1}"
        if fy not in out:
            out.append(fy)
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


def fy_label(fy):
    a, b = fy.split("-")
    return f"July {a} to June {b}"


# the employees grid: header has 'Employee Tin', and it is not the search form
GRID_JS = r"""(args) => {
  const norm = s => (s || '').replace(/[\u00a0\s]+/g, ' ').trim();
  const tick = args && args.tick;
  const limit = (args && args.limit) || 1e9;
  let ticked = 0;
  for (const t of document.querySelectorAll('table')) {
    if (t.querySelector('#period, #retBrnch') || t.querySelector('table')) continue;
    const head = t.querySelector('thead tr') || t.querySelector('tr');
    if (!head) continue;
    const heads = [...head.children].map(c => norm(c.innerText).toLowerCase());
    if (!heads.some(h => h.includes('employee tin') || h === 'tin')) continue;
    const tinCol = heads.findIndex(h => h.includes('tin'));
    const tccCol = heads.findIndex(h => h.includes('tcc'));
    const nameCol = heads.findIndex(h => h.includes('name'));
    const rows = [...t.querySelectorAll('tbody tr')].filter(r => r !== head && r.children.length >= 5);
    const out = [];
    rows.forEach((r, i) => {
      const cells = [...r.children];
      const val = c => { if (!c) return ''; const inp = c.querySelector('input[type=text]');
                         return norm(inp ? inp.value : c.innerText); };
      const tcc = tccCol >= 0 ? val(cells[tccCol]) : '';
      const cb = r.querySelector('input[type=checkbox]');
      const allow = r.querySelector('[id^=allowTCC]');
      const may = !tcc && !(allow && norm(allow.textContent) === 'N');
      if (tick && cb && !cb.disabled) {
        const want = may && ticked < limit;
        if (cb.checked && !want) cb.click();          // untick first, so the page's own counter stays right
      }
      if (tick && cb && !cb.disabled && may && ticked < limit && !cb.checked) cb.click();
      if (cb && cb.checked) ticked++;
      const link = [...r.querySelectorAll('a')].find(a => /month/i.test(a.innerText));
      if (link) link.setAttribute('data-bot-mw', String(i));
      out.push({i, tin: tinCol >= 0 ? val(cells[tinCol]) : '', name: nameCol >= 0 ? val(cells[nameCol]) : '',
                tcc, may, box: !!cb, checked: cb ? cb.checked : false, link: !!link});
    });
    return {found: true, heads, rows: out};
  }
  const body = norm(document.body ? document.body.innerText : '');
  return {found: false, none: /no\s+(record|data|result|detail|employee)s?|not\s+found/i.test(body)};
}"""

CLEAN_GRID_JS = r"""() => {
  for (const t of document.querySelectorAll('table')) {
    if (t.querySelector('#period, #retBrnch') || t.querySelector('table')) continue;
    const head = t.querySelector('thead tr') || t.querySelector('tr');
    if (!head || !/employee tin/i.test(head.innerText || '')) continue;
    const c = t.cloneNode(true);
    c.querySelectorAll('input[type=text]').forEach(i =>
        i.replaceWith(document.createTextNode(i.value || i.getAttribute('value') || '')));
    c.querySelectorAll('input[type=checkbox]').forEach(i => i.remove());
    return c.outerHTML;
  }
  return '';
}"""

MARK_JS = """() => { document.documentElement.setAttribute('data-bot-stale', '1');
  document.querySelectorAll('table, table *').forEach(e => e.setAttribute('data-bot-stale', '1'));
  document.querySelectorAll('body *').forEach(e => {          // old messages too
    if (e.children.length === 0 && /no\\s+(record|data|result|detail|employee)s?|not\\s+found|generated/i
        .test(e.textContent || '')) e.setAttribute('data-bot-stale', '1'); }); }"""
FRESH_JS = """() => {
  for (const t of document.querySelectorAll('table')) {
    if (t.querySelector('table')) continue;
    const fresh = !t.hasAttribute('data-bot-stale') || t.querySelector('tr:not([data-bot-stale])');
    if (fresh && /employee tin/i.test((t.querySelector('thead tr') || t.querySelector('tr') || {}).innerText || ''))
      return 'grid';
  }
  if (!document.querySelector('#TaxCreditCertForm, #period')) return '';   // not the PAYE page (8 Oct 2026)
  for (const e of document.querySelectorAll('body *:not([data-bot-stale])')) {
    if (e.children.length === 0 && /no\\s+(record|data|result|detail|employee)s?|not\\s+found/i
        .test(e.textContent || '')) return 'none';
  }
  if (!document.documentElement.hasAttribute('data-bot-stale') && document.readyState === 'complete')
    return 'reloaded';
  return '';
}"""
NONE_TEXT_JS = """() => { for (const e of document.querySelectorAll('body *:not([data-bot-stale])')) {
    const t = (e.textContent || '').replace(/\\s+/g, ' ').trim();
    if (e.children.length === 0 && /no\\s+(record|data|result|detail|employee)s?|not\\s+found/i.test(t))
      return t.slice(0, 150); } return ''; }"""
LIMIT_JS = "() => (typeof count === 'number' && count > 0) ? count : 5"


class SearchTimeout(RuntimeError):
    """No answer to Search in time (10.15: the year is tried once more, then recorded as failed)."""


class PayeBot:
    def __init__(self, page, log, check, debug_dir=None):
        self.page, self.log, self.check = page, log, check
        self.debug_dir = Path(debug_dir) if debug_dir else None
        self._searches_sent = []                 # times a Search request left the browser (10.15)
        try:
            page.on("request", self._on_request)
        except Exception:
            pass
        from engines.dialogs import guard_for
        self._guard = guard_for(page, log)       # one answer policy per page (finding H7)
        self.dialogs = self._guard.messages

    def _on_request(self, request):
        try:
            if "searchpayeentry" in request.url.lower():
                self._searches_sent.append(time.monotonic())
        except Exception:
            pass

    @staticmethod
    def _clear_guard(frame):
        """The PAYE page sends a form only once (frmSubmitFlag, copied from the live page); it is cleared when
        the page loads.  Pressing Search again on the same page needs it cleared, or the press does nothing."""
        try:
            frame.evaluate("() => { if (typeof frmSubmitFlag !== 'undefined') frmSubmitFlag = 0; }")
        except Exception:
            pass

    def _sent_since(self, t):
        return any(x >= t for x in self._searches_sent)

    def expecting_generate_confirm(self):
        return self._guard.expecting()

    # ---------------------------------------------------------------- basics
    def frames(self):
        try:
            return list(self.page.frames)
        except Exception:
            return [self.page.main_frame]

    def find(self, sels, timeout_s=0.0, visible=False):
        end = time.monotonic() + timeout_s
        while True:
            self.check()
            for fr in self.frames():
                for sel in sels:
                    try:
                        loc = fr.locator(sel).first
                        if loc.count() and (not visible or loc.is_visible()):
                            return fr, loc
                    except Exception:
                        continue
            if time.monotonic() >= end:
                return None, None
            self.page.wait_for_timeout(250)

    def close_processing(self):
        _, btn = self.find(SEL["processing_close"], 0, visible=True)
        if btn is not None:
            try:
                btn.evaluate("el => el.click()", timeout=3000)
                self.log("  closed the 'Processing' box.")
            except Exception:
                pass

    # ---------------------------------------------------------------- page
    def open_page(self):
        fr, per = self.find(SEL["period"], 0)
        if per is not None:
            return fr
        self.log(f"Opening '{MENU_ITEM}'...")
        end = time.monotonic() + 40
        while time.monotonic() < end:
            self.check()
            self.close_processing()
            for f in self.frames():
                try:
                    link = f.get_by_role("link", name=MENU_ITEM, exact=False).first
                    if not link.count():
                        link = f.locator(f"a:has-text({MENU_ITEM!r})").first
                    if link.count():
                        link.evaluate("el => el.click()", timeout=5000)
                        fr, per = self.find(SEL["period"], 25)
                        if per is not None:
                            self.log("PAYE certificate form is open.")
                            return fr
                except Exception:
                    continue
            self.page.wait_for_timeout(500)
        raise RuntimeError(f"could not open '{MENU_ITEM}' (menu link not found on the portal home page)")

    def search(self, fy):
        """Return For = Headquarter, Period = this financial year, Search; wait for NEW results."""
        fr = self.open_page()
        self.close_processing()
        br = fr.locator(SEL["branch"][0]).first
        if br.count():
            opts = br.evaluate("el => Array.from(el.options).map(o => [o.value, o.text.trim()])", timeout=5000)
            val = next((v for v, t in opts if v == "H"), None) or \
                next((v for v, t in opts if "headquarter" in t.lower()), None)
            if val:
                br.select_option(val, timeout=5000)
        per = fr.locator(SEL["period"][0]).first
        opts = per.evaluate("el => Array.from(el.options).map(o => [o.value, o.text.trim()])", timeout=5000)
        val = next((v for v, t in opts if v == fy), None) or \
            next((v for v, t in opts if t.lower() == fy_label(fy).lower()), None)
        if val is None:
            raise LookupError(f"the portal's Period list has no '{fy_label(fy)}'")
        per.select_option(val, timeout=5000)
        got = per.evaluate("el => el.options[el.selectedIndex].text", timeout=5000)
        self.log(f"Period: {got.strip()}")
        for f in self.frames():
            try:
                f.evaluate(MARK_JS)
            except Exception:
                pass
        bfr, btn = self.find([f"#TaxCreditCertForm {s}" if not s.startswith("#") else s
                              for s in ("button:has-text('Search')", "input[value='Search']")]
                             + SEL["search"], 5)
        if btn is None:
            raise RuntimeError("the Search button was not found")
        self._clear_guard(bfr)               # a search that got no answer left the page's guard set
        # the time BEFORE pressing: the browser may report the request while the press itself is still
        # returning (seen on the test portal - the search was sent, yet it was pressed a second time)
        t0 = time.monotonic()
        btn.evaluate("el => { setTimeout(() => el.click(), 0); }", timeout=5000)
        self.no_list_reason = ""
        end = t0 + WAIT_S
        reloaded_at = None
        pressed_again = False
        while time.monotonic() < end:
            self.check()
            if not pressed_again and time.monotonic() - t0 > PRESS_AGAIN_S and not self._sent_since(t0):
                # 10.15 (live 8 Oct 2026, 22:40: 60 s with no answer): the press never reached the portal.
                # The page's own guard (frmSubmitFlag) ignores a second press, so it is cleared first.
                pressed_again = True
                self.log(f"  Search sent nothing to the portal in {PRESS_AGAIN_S} s - pressing it again.")
                try:
                    self._clear_guard(bfr)
                    btn.evaluate("el => { setTimeout(() => el.click(), 0); }", timeout=5000)
                except Exception as e:
                    self.log(f"  Could not press Search again ({type(e).__name__}).")
            kinds = []
            for f in self.frames():
                try:
                    kinds.append((f, f.evaluate(FRESH_JS)))
                except Exception:
                    continue
            for f, kind in kinds:
                if kind == "grid":
                    self.page.wait_for_timeout(400)
                    self.log(f"  Employees list shown {time.monotonic() - t0:.0f} s after Search.")
                    return "grid"
            for f, kind in kinds:
                if kind == "none":
                    try:
                        said = f.evaluate(NONE_TEXT_JS)
                    except Exception:
                        said = ""
                    self.no_list_reason = f"the portal said '{said}'" if said else "the portal said there are no records"
                    self.evidence_frame = f
                    return "none"
            if reloaded_at is None and any(k == "reloaded" for _, k in kinds):
                reloaded_at = time.monotonic()          # the PAYE page itself came back: give a list time to draw
                self.evidence_frame = next(f for f, k in kinds if k == "reloaded")
            if reloaded_at is not None and time.monotonic() - reloaded_at > RELOAD_GRACE_S:
                self.no_list_reason = (f"the PAYE page came back {reloaded_at - t0:.0f} s after Search with no "
                                       f"employees list and no message (waited {RELOAD_GRACE_S} s more)")
                return "none"
            self.page.wait_for_timeout(300)
        sent = "the portal received the search" if self._sent_since(t0) else "the search never reached the portal"
        raise SearchTimeout(f"no results within {WAIT_S} s after Search ({sent})")

    def save_evidence(self, stem):
        """A copy of the PAYE page and a screenshot, so a 'no employees' answer can be checked later."""
        if self.debug_dir is None:
            return []
        saved = []
        try:
            self.debug_dir.mkdir(parents=True, exist_ok=True)
            f = getattr(self, "evidence_frame", None) or self.page.main_frame
            path = self.debug_dir / f"{stem}.html"
            path.write_text(f.content(), encoding="utf-8")
            saved.append(path)
        except Exception:
            pass
        try:
            path = self.debug_dir / f"{stem}.png"
            self.page.screenshot(path=str(path), full_page=True)
            saved.append(path)
        except Exception:
            pass
        return saved

    def grid(self, tick=False):
        for f in self.frames():
            try:
                res = f.evaluate(GRID_JS, {"tick": tick})
            except Exception:
                continue
            if res.get("found"):
                res["frame"] = f
                return res
        return {"found": False, "rows": []}

    def grid_frame(self):
        import pandas as pd
        for f in self.frames():
            try:
                html = f.evaluate(CLEAN_GRID_JS)
            except Exception:
                continue
            if html:
                df = pd.read_html(io.StringIO(html))[0]
                df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
                def is_link_col(col):
                    vals = [str(v) for v in df[col].tolist() if not pd.isna(v)]
                    return bool(vals) and all("month-wise" in v.lower() for v in vals)
                keep = [c for c in df.columns
                        if not c.lower().startswith("select") and not is_link_col(c)]
                df = df[keep]
                from engines.wht_outputs import ID_COL, _id_text
                for c in df.columns:                 # TINs, Sr No., TCC No. stay text
                    if ID_COL.search(str(c)):
                        df[c] = df[c].map(_id_text)
                return df.reset_index(drop=True)
        return None

    # ---------------------------------------------------------------- generate
    def generate_missing(self, fy=None):
        """Press GenerateTCC for the employees without a TCC No., in batches of the page's limit (live: 5),
        searching again after each batch so the next batch sees the new numbers.  Returns how many were
        sent.  Stops when none are left, or when a batch did not reduce the number missing."""
        total, rounds = 0, 0
        while True:
            g = self.grid()
            missing = [r for r in g.get("rows", []) if r.get("may", not r["tcc"])]
            if not missing:
                return total
            if rounds and len(missing) >= before:
                self.log(f"  {len(missing)} employee(s) still without a TCC No. after the last batch - stopping.")
                return total
            before = len(missing)
            limit = 5
            try:
                limit = int(g["frame"].evaluate(LIMIT_JS))
            except Exception:
                pass
            rounds += 1
            total += self._generate_batch(missing, limit, rounds)
            if fy is None:
                return total
            self.search(fy)                              # the new TCC numbers show only now

    def _generate_batch(self, missing, limit, rounds):
        g = self.grid(tick=True) if limit >= 10 ** 6 else self._tick(limit)
        ticked = [r for r in g["rows"] if r["checked"]]
        if not ticked:
            raise RuntimeError(f"{len(missing)} employee(s) have no TCC No. but none could be ticked")
        self.log(f"  Batch {rounds}: {len(ticked)} of {len(missing)} employee(s) without a TCC No. ticked "
                 f"(the portal takes at most {limit} at a time) - pressing GenerateTCC...")
        _, btn = self.find(SEL["generate"], 5)
        if btn is None:
            raise RuntimeError("the GenerateTCC button was not found")
        self.dialogs.clear()
        for f in self.frames():
            try:
                f.evaluate(MARK_JS)
            except Exception:
                pass
        with self.expecting_generate_confirm():
            btn.evaluate("el => { setTimeout(() => el.click(), 0); }", timeout=5000)
            self._wait_generate_answer()
        for m in self.dialogs:
            self.log(f"  Portal message: {m}")
        self.page.wait_for_timeout(800)
        return len(ticked)

    def _tick(self, limit):
        for f in self.frames():
            try:
                res = f.evaluate(GRID_JS, {"tick": True, "limit": limit})
            except Exception:
                continue
            if res.get("found"):
                res["frame"] = f
                return res
        return {"found": False, "rows": []}

    def _find_monthwise(self, before):
        return _find_monthwise_impl(self, before)

    @staticmethod
    def _close_window(win):
        _close_window(win)

    def _wait_generate_answer(self):
        end = time.monotonic() + WAIT_S
        while time.monotonic() < end:            # wait for the portal to answer
            self.check()
            if any("select" in m.lower() for m in self.dialogs):
                raise RuntimeError(self.dialogs[-1])
            answered = False
            for f in self.frames():
                try:
                    if f.evaluate(FRESH_JS) or f.evaluate(
                            "() => /generated|success/i.test(document.body.innerText)"):
                        answered = True
                        break
                except Exception:
                    continue
            if answered:
                break
            self.page.wait_for_timeout(400)

    # ---------------------------------------------------------------- month-wise details
    def month_wise(self, row):
        """Open one employee's 'Get Month-Wise Details', read the table, Close."""
        import pandas as pd
        self.check()
        g = self.grid()
        f = g.get("frame")
        if f is None:
            raise RuntimeError("the employees list is not on screen")
        link = f.locator(f"[data-bot-mw='{row['i']}']").first
        if not link.count():
            raise RuntimeError("no 'Get Month-Wise Details' link in this row")
        context = self.page.context
        before = set(context.pages)
        link.evaluate("el => { setTimeout(() => el.click(), 0); }", timeout=5000)
        win, popup = self._find_monthwise(before)
        if popup is None:
            for w in [p for p in context.pages if p not in before]:
                self._close_window(w)
            raise RuntimeError(f"the month-wise details did not open within {MONTHWISE_WAIT_S} s")
        self.page.wait_for_timeout(400 if win is None else 100)
        html = popup.evaluate("el => el.outerHTML", timeout=5000)
        if self.debug_dir is not None and not getattr(self, "_saved_example", False):
            self.debug_dir.mkdir(parents=True, exist_ok=True)
            (self.debug_dir / "paye_monthwise_example.html").write_text(html, encoding="utf-8")
            self._saved_example = True
        try:
            tables = pd.read_html(io.StringIO(html))
        except ValueError:
            tables = []
        best = None
        for t in tables:
            txt = " ".join("" if pd.isna(v) else str(v) for v in t.values.ravel()).lower()
            score = sum(m in txt for m in ("jan", "feb", "mar", "apr", "may", "jun", "jul",
                                           "aug", "sep", "oct", "nov", "dec")) * 100 + t.size
            if best is None or score > best[0]:
                best = (score, t)
        if win is not None:                          # live portal: its own window - closed, like Close does
            for w in [p for p in context.pages if p not in before]:
                self._close_window(w)
        else:
            _, close = self.find(SEL["popup_close"], 5)
            if close is not None:
                close.evaluate("el => el.click()", timeout=5000)
            for _ in range(20):
                _, still = self.find(SEL["popup"], 0, visible=True)
                if still is None:
                    break
                self.page.wait_for_timeout(250)
        if best is None:
            raise RuntimeError("the month-wise details box held no table")
        df = best[1].copy()
        df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
        df = df.dropna(axis=0, how="all")
        df = df[[not any(str(v).strip().lower() == "close" for v in row if not pd.isna(v))
                 for row in df.values]]
        for c in df.columns:
            if "month" not in c.lower():
                continue
            df[c] = df[c].map(lambda v: "" if pd.isna(v) else str(v))
        df.insert(0, "Employee Name", row["name"])
        df.insert(0, "Employee TIN", row["tin"])
        return df.reset_index(drop=True)


def _close_window(win):
    try:
        if not win.is_closed():
            win.close()
    except Exception:
        pass


def _find_monthwise_impl(bot, before):
    """(window or None, #divPostContent) - in a new pop-up window (live portal) or in the portal window."""
    end = time.monotonic() + MONTHWISE_WAIT_S
    while time.monotonic() < end:
        bot.check()
        for win in [p for p in bot.page.context.pages if p not in before]:
            try:
                if win.is_closed():
                    continue
                for fr in win.frames:
                    box = fr.locator(SEL["popup"][0]).first
                    if box.count() and box.is_visible():
                        return win, box
            except Exception:
                continue
        _, box = bot.find(SEL["popup"], 0, visible=True)
        if box is not None:
            return None, box
        bot.page.wait_for_timeout(200)
    return None, None


def _money_cols(df):
    from engines.wht_outputs import ID_COL, as_number
    out = []
    for c in df.columns:
        if ID_COL.search(str(c)) or "name" in str(c).lower() or "month" in str(c).lower():
            continue
        vals = [v for v in df[c].tolist() if str(v).strip() not in ("", "nan")]
        if vals and sum(as_number(v) is not None for v in vals) >= 0.6 * len(vals):
            out.append(c)
    return out


def save_workbook(path, tin, fy, employees, monthwise):
    from openpyxl import Workbook
    from engines.wht_outputs import _write_sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Employees"
    _write_sheet(ws, employees, f"PAYE tax credit certificates - {fy_label(fy)}",
                 subtitle=f"Employer TIN {tin}", amount_cols=_money_cols(employees))
    if monthwise is not None and len(monthwise):
        ws2 = wb.create_sheet("Month-wise details")
        _write_sheet(ws2, monthwise, f"Month-wise details - {fy_label(fy)}",
                     subtitle="one block per employee, as shown by 'Get Month-Wise Details'",
                     amount_cols=_money_cols(monthwise), totals=False)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return str(path)


def run(page, tin, period_from, period_to, out_dir, log, check, record,
        generate=True, monthwise=True, debug_dir=None, reopen=None, failed=None, years=None):
    """All financial years touched by the period (or `years`). Returns the list of files saved.
    10.15: a Search with no answer keeps a page copy, is tried once more (after `reopen()`, which brings back
    the portal home page, when given) and then recorded as failed for THAT year only - the next year goes on.
    Failed years are added to `failed` (for the second try at the end of the client)."""
    import pandas as pd
    bot = PayeBot(page, log, check, debug_dir=debug_dir)
    years = list(years) if years else financial_years(period_from, period_to)
    log(f"PAYE certificates: period {period_from} to {period_to} -> "
        + ", ".join(fy_label(y) for y in years))
    files = []
    for fy in years:
        check()
        sec = f"PAYE certificates {fy_label(fy)}"
        try:
            kind = _search_twice(bot, fy, log, reopen)
        except LookupError as e:
            log(f"  {e} - skipped.")
            record(sec, "skipped", str(e))
            continue
        except SearchTimeout as e:
            record(sec, "failed", str(e))
            if failed is not None:
                failed.append(fy)
            continue
        if kind == "none" or not bot.grid().get("rows"):
            why = getattr(bot, "no_list_reason", "") or "no employees list on the page"
            saved = bot.save_evidence(f"paye_no_employees_{fy}")
            log(f"  No employees listed for {fy_label(fy)} - {why}."
                + (f" Page copy: {saved[0].name}" if saved else ""))
            record(sec, "no records", f"No employees list: {why}")
            continue
        made = 0
        if generate:
            try:
                made = bot.generate_missing(fy)
            except Exception as e:
                msg = (str(e).splitlines() or [""])[0][:150]
                log(f"  WARNING: TCC numbers not generated: {msg}")
                record(f"{sec} - generate TCC", "failed", msg)
        g = bot.grid()
        left = [r for r in g["rows"] if not r["tcc"]]
        log(f"  {len(g['rows'])} employee(s); Generate pressed for {made}; "
            f"{len(left)} still without a TCC No.")
        if left:
            record(f"{sec} - TCC numbers", "incomplete", f"{len(left)} employee(s) without a TCC No.")
        employees = bot.grid_frame()
        details = []
        if monthwise:
            for row in g["rows"]:
                if not row["link"]:
                    continue
                try:
                    df = bot.month_wise(row)
                    details.append(df)
                    log(f"  Month-wise details: {row['name'] or row['tin']} ({len(df)} line(s))")
                except Exception as e:
                    msg = (str(e).splitlines() or [""])[0][:150]
                    log(f"  WARNING: month-wise details for {row['name'] or row['tin']}: {msg}")
                    record(f"{sec} - month-wise {row['tin']}", "failed", msg)
                    if not bot.grid().get("found"):   # list gone after the box closed: search again
                        _search_twice(bot, fy, log, reopen)
        mw = pd.concat(details, ignore_index=True) if details else None
        path = save_workbook(Path(out_dir) / f"PAYE_TCC_{tin}_{fy}.xlsx", tin, fy,
                             employees if employees is not None else pd.DataFrame(), mw)
        log(f"Saved: {path}")
        record(f"{sec} - workbook", "saved", f"{len(g['rows'])} employee(s)", path)
        files.append(path)
    return files


def _search_twice(bot, fy, log, reopen=None):
    """bot.search(fy); on no answer: a page copy, then once more (after reopen() when given)."""
    try:
        return bot.search(fy)
    except SearchTimeout as e:
        saved = bot.save_evidence(f"paye_search_timeout_{fy}")
        log(f"  {e}." + (f" Page copy: {', '.join(p.name for p in saved)}." if saved else "")
            + " Trying once more" + (" from the portal home page." if reopen else "."))
    if reopen is not None:
        reopen()
    try:
        return bot.search(fy)
    except SearchTimeout as e:
        saved = bot.save_evidence(f"paye_search_timeout_{fy}_2")
        log(f"  {e} - {fy_label(fy)} recorded as failed; the other years go on."
            + (f" Page copy: {', '.join(p.name for p in saved)}." if saved else ""))
        raise
