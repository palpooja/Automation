"""
modules/returns.py - Return history -> PDF / .mhtml + Returns_Status.xlsx

Browser work: engines/returns_pw.py (Playwright port of the former
ura_returns_bot.pyw v4.3; the Selenium-free helpers are in returns_common.py
and the status-sheet writer is reused).
"""
import tkinter as tk
from tkinter import ttk

from core.runner import Module
from core.widgets import WebButton
from engines import returns_common as eng


class _Tip:
    """Hover text for a shortened label (shows the full return name)."""
    def __init__(self, widget, text):
        self.w, self.text, self.tip = widget, text, None
        widget.bind("<Enter>", self.show); widget.bind("<Leave>", self.hide)

    def show(self, _e=None):
        self.tip = tk.Toplevel(self.w); self.tip.wm_overrideredirect(True)
        x, y = self.w.winfo_rootx(), self.w.winfo_rooty() + self.w.winfo_height() + 2
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, bg="#FFFBE6", relief="solid", bd=1,
                 padx=6, pady=3).pack()

    def hide(self, _e=None):
        if self.tip:
            self.tip.destroy(); self.tip = None


class ReturnsModule(Module):
    id = "returns"
    title = "Returns"
    shared_period = None          # (from 'YYYY-MM', to 'YYYY-MM'), set by the window
    description = ("Copies of every return filed in the period: PDF with schedules (Local Excise Duty "
                   "as a web-page copy) plus Returns_Status.xlsx")

    def build_panel(self, parent):
        f = ttk.Frame(parent, padding=10)
        ttk.Label(f, text=self.title, style="Head.TLabel").grid(row=0, column=0,
                                                                 columnspan=2, sticky="w")
        ttk.Label(f, text=self.description, style="Muted.TLabel").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(0, 8))

        pf = ttk.LabelFrame(f, text=" Options ", padding=8)
        pf.grid(row=2, column=0, sticky="nw")
        self.patient = tk.BooleanVar(value=True)
        self.skip_existing = tk.BooleanVar(value=True)
        ttk.Checkbutton(pf, text="Wait until each return downloads (never skip a slow one)",
                        variable=self.patient).grid(row=0, column=0, sticky="w", padx=(0, 16))
        ttk.Checkbutton(pf, text="Skip returns already saved (judged by return number)",
                        variable=self.skip_existing).grid(row=0, column=1, sticky="w")
        cf = ttk.LabelFrame(f, text=" Return types ", padding=8)
        cf.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        self.cat_vars = {}
        ncol = 3
        rows = (len(eng.ALL_CATEGORIES) + ncol - 1) // ncol
        for i, cat in enumerate(eng.ALL_CATEGORIES):
            v = tk.BooleanVar(value=cat in eng.DEFAULT_TICKED)
            self.cat_vars[cat] = v
            short = cat if len(cat) <= 40 else cat[:38] + "…"
            cb = ttk.Checkbutton(cf, text=short, variable=v)
            cb.grid(row=i % rows, column=i // rows, sticky="w", padx=(0, 10))
            if short != cat:
                _Tip(cb, cat)
        bb = ttk.Frame(cf)
        bb.grid(row=rows, column=0, columnspan=ncol, sticky="w", pady=(6, 0))
        WebButton(bb, text="Select all", command=lambda: self._set_all(True)).pack(side="left")
        WebButton(bb, text="Clear all", command=lambda: self._set_all(False)).pack(side="left", padx=6)
        return f

    def _set_all(self, val):
        for v in self.cat_vars.values():
            v.set(val)

    def collect(self):
        if not self.shared_period:
            raise ValueError("Returns: set the period first.")
        import calendar
        pf, pt = self.shared_period
        y1, m1 = (int(x) for x in pf.split("-"))
        y2, m2 = (int(x) for x in pt.split("-"))
        d_from = f"01/{m1:02d}/{y1}"
        d_to = f"{calendar.monthrange(y2, m2)[1]:02d}/{m2:02d}/{y2}"
        ticked = [c for c, v in self.cat_vars.items() if v.get()]
        if not ticked:
            raise ValueError("Returns: tick at least one return type.")
        wanted = None if len(ticked) == len(eng.ALL_CATEGORIES) else ticked
        return {"date_from": d_from, "date_to": d_to, "wanted": wanted,
                "patient": self.patient.get(), "skip_existing": self.skip_existing.get()}

    def run(self, ctx, client, options, page=None):
        from engines import returns_pw
        if page is None:
            raise RuntimeError("Returns needs a logged-in session (core.session.open_session).")
        result = {}

        def log(msg):
            low = msg.lower()
            kind = ("err" if "error" in low or "***" in msg else "ok" if "saved" in low else "")
            ctx.log(msg, kind)
        try:
            returns_pw.run_on_page(page, client["login"], options["date_from"], options["date_to"],
                                   options["wanted"], log, ctx.check,
                                   str(ctx.output_root / "Returns"), patient=options["patient"],
                                   result=result, skip_existing=options.get("skip_existing", True),
                                   first_try_s=None if ctx.final_try else eng.FIRST_TRY_SECONDS,
                                   person=ctx if ctx.final_try else None)
        finally:                    # report work done so far, also after Stop
            for it in result.get("items", []):
                ctx.record(it.get("item", ""), it.get("status", ""), it.get("detail", ""),
                           it.get("file", ""))
        aside = result.get("deferred", [])
        failed_types = self.failed_types(result.get("items", []))
        if (aside or failed_types) and not ctx.final_try:   # 10.14 / 10.15: once more at the end of the client
            if aside:
                log(f"{len(aside)} return(s) set aside because the portal was slow to open them - they are "
                    "tried again with the full wait after this client's other steps.")
            if failed_types:
                log(f"Returns not downloaded ({', '.join(failed_types)}) are tried again after this client's "
                    "other steps.")
            what = " and ".join(x for x in (f"{len(aside)} slow return(s)" if aside else "",
                                            f"{len(failed_types)} return type(s) with failures" if failed_types
                                            else "") if x)
            ctx.after_client.append((f"Returns - second try for {what}",
                                     lambda pg: self.second_try(ctx, client, options, pg, aside, failed_types)))
        files = sum(n for _, n in result.get("summary", []))
        return {"files": files, "ok": result.get("ok", False), "folder": result.get("folder", "")}

    @staticmethod
    def failed_types(items):
        """Return types with a failure the second try can fix (10.15): a return not downloaded after every
        retry, or a type whose search failed.  Not the ones the person chose to skip."""
        out = []
        for it in items:
            if it.get("status") != "failed" or "skipped by you" in str(it.get("detail", "")):
                continue
            kind = str(it.get("item", "")).split(": ", 1)[0]
            if kind and kind not in out:
                out.append(kind)
        return out

    def second_try(self, ctx, client, options, page, aside, failed_types=()):
        """The returns set aside, with the mode's full wait; after a minute the person hears 'Attention
        needed' and may press Skip (the wait goes on while the question shows).  10.15: also every return
        type with a failure - returns already on disk are skipped, so only the missing ones are fetched."""
        from engines import returns_pw
        labels = {a["label"] for a in aside}
        types = sorted({a["type"] for a in aside} | set(failed_types))
        result = {}

        def log(msg):
            low = msg.lower()
            ctx.log(msg, "err" if "error" in low or "***" in msg else "ok" if "saved" in low else "")
        try:
            returns_pw.run_on_page(page, client["login"], options["date_from"], options["date_to"],
                                   types, log, ctx.check,
                                   str(ctx.output_root / "Returns"), patient=options["patient"],
                                   result=result, skip_existing=True, first_try_s=None, person=ctx)
        finally:
            got = set()
            for it in result.get("items", []):
                item = it.get("item", "")
                hit = next((lb for lb in labels if item.endswith(f": {lb}")), None)
                if hit and it.get("status") == "saved":
                    got.add(hit)
                # final_try: an item of the first try is updated ('saved on second try'), not listed twice
                ctx.record(item, it.get("status", ""), it.get("detail", ""), it.get("file", ""))
            for a in aside:
                if a["label"] in got:
                    ctx.log(f"Second try: {a['type']}: {a['label']} saved.", "ok")
        return {"files": len(got)}
