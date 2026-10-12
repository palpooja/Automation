"""
modules/ledger.py - Tax Statement ledgers (Level 2 / Level 3), tax-credit
certificates, taxpayer profile and TIN certificate.

The browser work is engines/ledger_engine.py (the former ura_bot.pyw,
unchanged apart from the output-folder switch).  This file is only the
options page and the glue.
"""
import contextlib
import datetime
import tkinter as tk
from tkinter import ttk

from core.runner import Module, PrintToLog, Stopped
from engines import ledger_engine as eng


class LedgerModule(Module):
    id = "ledger"
    title = "Ledgers"
    shared_period = None          # (from 'YYYY-MM', to 'YYYY-MM'), set by the window
    description = ("Tax statement ledgers month by month (Level 2 and Level 3), taxpayer profile "
                   "and TIN (Tax Identification Number) certificate")

    TAX_LABELS = {
        "Value Added Tax": "VAT (Value Added Tax)",
        "IT-WHT": "IT-WHT (Income Tax - Withholding Tax)",
        "PAYE": "PAYE (Pay As You Earn - tax on salaries)",
        "Local Excise Duty": "Local Excise Duty",
        "Income Tax": "Income Tax (yearly)",
        "Rental": "Rental Income Tax (yearly)",
    }
    OTHER_LABELS = {
        "profile": "Taxpayer profile (registration details)",
        "tin_cert": "TIN / VAT certificate",
        "wht_tcc": "Withholding tax credit certificates",
        "paye_tcc": "PAYE (Pay As You Earn) tax credit certificates",
    }

    def build_panel(self, parent):
        f = ttk.Frame(parent, padding=10)
        ttk.Label(f, text=self.title, style="Head.TLabel").grid(row=0, column=0,
                                                                 columnspan=3, sticky="w")
        ttk.Label(f, text=self.description, style="Muted.TLabel").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(0, 8))

        # tax types
        tf = ttk.LabelFrame(f, text=" Ledgers to download ", padding=8)
        tf.grid(row=2, column=0, sticky="nsew", padx=(0, 8))
        self.tax_vars = {}
        for t in eng.TAX_TYPES + eng.ANNUAL_TAX_TYPES:
            v = tk.BooleanVar(value=t in eng.TAX_TYPES)
            self.tax_vars[t] = v
            ttk.Checkbutton(tf, text=self.TAX_LABELS.get(t, t), variable=v).pack(anchor="w")

        # other documents
        of = ttk.LabelFrame(f, text=" Other documents ", padding=8)
        of.grid(row=2, column=1, sticky="nsew")
        self.other_vars = {}
        for key, label in eng.OTHER_JOBS.items():
            if key in ("wht_tcc", "paye_tcc"):  # step 3 Certificates does these
                continue
            v = tk.BooleanVar(value=False)
            self.other_vars[key] = v
            ttk.Checkbutton(of, text=self.OTHER_LABELS.get(key, label), variable=v).pack(anchor="w")
        ttk.Label(of, text="Withholding and PAYE tax credit certificates: see step 3, "
                           "Certificates.",
                  style="Muted.TLabel", justify="left", wraplength=300).pack(anchor="w", pady=(6, 0))

        # options - full sentences, dependent option indented under its parent
        opt = ttk.LabelFrame(f, text=" Options ", padding=8)
        opt.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Label(opt, text="Each ledger is saved at Level 2 (the summary for each month) and Level 3 "
                            "(the transactions behind each month). Income Tax and Rental Income Tax "
                            "are yearly: they get Level 2 only, unless you tick the second option.",
                  style="Muted.TLabel", wraplength=900, justify="left").grid(
            row=0, column=0, sticky="w", pady=(0, 4))
        self.skip_nil = tk.BooleanVar(value=True)
        self.annual_l3 = tk.BooleanVar(value=False)
        self.annual_all = tk.BooleanVar(value=False)
        self.speed = tk.BooleanVar(value=True)
        self.debug = tk.BooleanVar(value=False)
        self.refresh_l2 = tk.BooleanVar(value=True)
        self.keep_months = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="Skip the Level 3 transactions for months where nothing was paid",
                        variable=self.skip_nil).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(opt, text="Income Tax and Rental Income Tax: also download their Level 3 transactions",
                        variable=self.annual_l3, command=self._sync_annual).grid(row=2, column=0, sticky="w")
        self.annual_all_cb = ttk.Checkbutton(
            opt, text="Include every year in the ledger, not only the period chosen above",
            variable=self.annual_all)
        self.annual_all_cb.grid(row=3, column=0, sticky="w", padx=(26, 0))
        ttk.Checkbutton(opt, text="Load portal pages faster (skips the chat window and pictures from "
                                  "other websites)", variable=self.speed).grid(row=4, column=0, sticky="w")
        ttk.Checkbutton(opt, text="Save a screenshot at every step, for troubleshooting (slower)",
                        variable=self.debug).grid(row=5, column=0, sticky="w")
        ttk.Checkbutton(opt, text="Fetch a fresh Level 2 ledger if the saved one is from an earlier day "
                                  "(the old copy is kept in 'previous')",
                        variable=self.refresh_l2).grid(row=6, column=0, sticky="w")
        ttk.Checkbutton(opt, text="Also keep a separate Excel file for every Level 3 month (each month is "
                                  "already a tab in the tax type's workbook)",
                        variable=self.keep_months).grid(row=7, column=0, sticky="w")
        # 10.16: how Level 3 is fetched (user, 9 Oct 2026: 'build it, then we decide how to implement it')
        rf = ttk.Frame(opt)
        rf.grid(row=8, column=0, sticky="w", pady=(6, 0))
        ttk.Label(rf, text="How Level 3 is fetched:").pack(side="left")
        self.l3_route = tk.StringVar(value=self.ROUTES[0][1])
        ttk.Combobox(rf, textvariable=self.l3_route, values=[t for _, t in self.ROUTES], state="readonly",
                     width=62).pack(side="left", padx=(6, 0))
        sf = ttk.Frame(opt)
        sf.grid(row=9, column=0, sticky="w", pady=(4, 0))
        ttk.Label(sf, text="Switch to the screen way if URA has not answered after").pack(side="left")
        self.direct_switch = tk.StringVar(value="0")
        ttk.Spinbox(sf, from_=0, to=600, increment=5, textvariable=self.direct_switch, width=5).pack(
            side="left", padx=4)
        ttk.Label(sf, text="seconds (0 = never: wait for URA's answer)", style="Muted.TLabel").pack(side="left")
        self._sync_annual()
        return f

    ROUTES = [("auto", "Auto - direct request, proved on the first month of each tax type"),
              ("direct", "Direct request first, no proof (screen way if it fails)"),
              ("screen", "Screen way only - click, Export To Excel (as version 10.15)")]

    def _sync_annual(self):
        """'every year' only means something when the yearly Level 3 option is on."""
        if self.annual_l3.get():
            self.annual_all_cb.state(["!disabled"])
        else:
            self.annual_all.set(False)
            self.annual_all_cb.state(["disabled"])

    def collect(self):
        if not self.shared_period:
            raise ValueError("Ledgers: set the period first.")
        pf, pt = self.shared_period
        try:
            a = datetime.datetime.strptime(pf, "%Y-%m")
            b = datetime.datetime.strptime(pt, "%Y-%m")
        except ValueError:
            raise ValueError("Ledgers: type the period as YYYY-MM, e.g. 2025-07.")
        if a > b:
            raise ValueError("Ledgers: 'From' is after 'To'.")
        taxes = [t for t, v in self.tax_vars.items() if v.get()]
        others = [k for k, v in self.other_vars.items() if v.get()]
        if not taxes and not others:
            raise ValueError("Ledgers: tick at least one ledger or document.")
        return {"period_from": pf, "period_to": pt, "tax_types": taxes,
                "other_jobs": others, "skip_nil": self.skip_nil.get(),
                "annual_l3": self.annual_l3.get(),
                "annual_all_years": self.annual_all.get(),
                "speed": self.speed.get(), "debug": self.debug.get(),
                "refresh_l2": self.refresh_l2.get(),
                "keep_month_files": self.keep_months.get(),
                "l3_route": next((k for k, t in self.ROUTES if t == self.l3_route.get()), "auto"),
                "direct_switch_s": self._switch_seconds()}

    def _switch_seconds(self):
        try:
            return max(0, int(float(self.direct_switch.get() or 0)))
        except ValueError:
            raise ValueError("Ledgers: the switch time must be a number of seconds (0 = never).")

    def run(self, ctx, client, options, page=None):
        """`page` = an already-logged-in page from core.session; without it
        the engine starts its own browser (stand-alone use)."""
        eng.OUTPUT_ROOT = str(ctx.output_root / "Ledgers")
        eng.WORK_DIR = str(ctx.log_path.parent)
        eng.SPEED_UPS = options["speed"]
        eng.DEBUG_DUMPS = eng.DEBUG_BLOCKING = options["debug"]
        eng.STOP_EVENT.clear()
        eng.RESULTS.clear()                  # never report the previous client's items
        eng.PROMPT_HANDLER = lambda message, kind, defaults, poll: ctx.ask(message, kind, defaults, poll)
        # PAYE certificates are step 3 (Certificates); this page never offers them
        settings = dict(options, tin=client["login"], password=client["password"],
                        other_jobs=[j for j in options["other_jobs"] if j not in ("paye_tcc", "wht_tcc")])
        has_engine_work = bool(settings["tax_types"] or settings["other_jobs"])
        bridge = _StopBridge(ctx, eng.STOP_EVENT)
        bridge.start()
        try:
            with contextlib.redirect_stdout(PrintToLog(ctx)), \
                 contextlib.redirect_stderr(PrintToLog(ctx)):
                try:
                    failed = []
                    if has_engine_work:
                        failed = (eng.run_job_on_page(page, settings) if page is not None
                                  else eng.run_job(settings))
                except eng.StopRequested:
                    raise Stopped()
        finally:
            bridge.stop()
            for r in eng.RESULTS:           # report work done so far, also after Stop
                ctx.record(f"{r.get('Section', '')}: {r.get('Item', '')}", r.get("Status", ""),
                           r.get("Detail", ""), r.get("File", ""))
        if ctx.stop_event.is_set():
            raise Stopped()
        taxes, others = self.failed_work(eng.RESULTS, settings)
        if (taxes or others) and not ctx.final_try:     # 10.15: once more at the end of the client
            what = ", ".join(taxes + [self.OTHER_NAMES.get(o, o) for o in others])
            ctx.log(f"Ledgers not finished ({what}) - tried again after this client's other steps.")
            ctx.after_client.append((f"Ledgers - second try for {what}",
                                     lambda pg: self.second_try(ctx, client, options, pg, taxes, others)))
        return {"files": len([r for r in eng.RESULTS if r.get("File")]),
                "failed": len(failed)}

    OTHER_NAMES = {"profile": "taxpayer profile", "tin_cert": "TIN/VAT certificate"}

    @classmethod
    def failed_work(cls, results, settings):
        """(tax types, other jobs) with a failed item: a whole tax type, a month or year of it, a profile."""
        taxes, others = [], []
        by_section = {f"{t} ledger": t for t in settings.get("tax_types", [])}
        by_label = {v: k for k, v in cls.OTHER_NAMES.items()}
        for r in results:
            if "fail" not in str(r.get("Status", "")).lower():
                continue
            sec = r.get("Section", "")
            if sec in by_section and by_section[sec] not in taxes:
                taxes.append(by_section[sec])
            elif sec in by_label and by_label[sec] in settings.get("other_jobs", []) \
                    and by_label[sec] not in others:
                others.append(by_label[sec])
        return taxes, others

    def second_try(self, ctx, client, options, page, taxes, others):
        """Only the tax types / documents with a failure; months and years already saved are skipped by the
        engine's resume rules, so only what is missing is fetched.  The first try's items are updated."""
        res = self.run(ctx, client, dict(options, tax_types=taxes, other_jobs=others), page)
        still = {r.get("Section") for r in eng.RESULTS if "fail" in str(r.get("Status", "")).lower()}
        for tax in taxes:
            if f"{tax} ledger" not in still:
                ctx.resolve(f"{tax} ledger: whole tax type", "done on second try")
        return res


class _StopBridge:
    """Mirrors the suite's stop flag onto the engine's own STOP_EVENT."""
    import threading as _th

    def __init__(self, ctx, engine_event):
        self.ctx, self.ev = ctx, engine_event
        self._done = self._th.Event()
        self._t = self._th.Thread(target=self._loop, daemon=True)

    def start(self):
        self._t.start()

    def stop(self):
        self._done.set()

    def _loop(self):
        while not self._done.is_set():
            if self.ctx.stop_event.is_set():
                self.ev.set()
            self._done.wait(0.3)
