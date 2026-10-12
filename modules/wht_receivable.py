"""
modules/wht_receivable.py - month-by-month 'Generate Tax Credit Certificate
for Withholding' search -> Excel / PDF / CSV / page copies.

Browser work: engines/wht_engine.py (WhtBot + run_months, rebuilt from
ura_bot.pyc).  Login uses the Ledger engine's login() so the suite's
browser / proxy / timeout settings and the yellow bar apply.
Files: engines/wht_outputs.py.
"""
import datetime
import os
import tkinter as tk
from tkinter import ttk

from core.runner import Module, Stopped


def uganda_fy(today=None):
    """Uganda's financial year runs 1 July - 30 June."""
    today = today or datetime.date.today()
    start_year = today.year if today.month >= 7 else today.year - 1
    return (start_year, 7), (start_year + 1, 6)


class WhtReceivableModule(Module):
    """Step 3 of the client briefing: withholding and PAYE tax credit certificates.
    (File name kept for compatibility; the menu calls it 'Certificates'.)"""
    id = "certificates"
    title = "Certificates"
    description = ("Withholding tax (WHT) and PAYE (Pay As You Earn) tax credit certificates, "
                   "saved to Excel")
    shared_period = None          # (from 'YYYY-MM', to 'YYYY-MM'), set by the window
    # Generating a TCC creates a certificate on the client's URA account, so it
    # happens only when the person ticks it for this run (finding H7).
    DEFAULT_GENERATE_WHT_TCC = False
    DEFAULT_GENERATE_PAYE_TCC = False

    def build_panel(self, parent):
        f = ttk.Frame(parent, padding=10)
        ttk.Label(f, text=self.title, style="Head.TLabel").grid(row=0, column=0,
                                                                 columnspan=3, sticky="w")
        ttk.Label(f, text=self.description, style="Muted.TLabel").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(0, 8))

        wf = ttk.LabelFrame(f, text=" Withholding tax (WHT) ", padding=8)
        wf.grid(row=2, column=0, sticky="nsew", padx=(0, 8))
        self.gen_vars = {}
        for code, label, tip, on in (
                ("Payee", "Payee (Deductee)", "tax others withheld from this client", True),
                ("Payer", "Withholding Agent (Payer)", "tax this client withheld from others", False)):
            v = tk.BooleanVar(value=on)
            self.gen_vars[code] = v
            ttk.Checkbutton(wf, text=label, variable=v).pack(anchor="w")
            ttk.Label(wf, text=tip, style="Muted.TLabel").pack(anchor="w", padx=(22, 0))
        self.gen_tcc = tk.BooleanVar(value=self.DEFAULT_GENERATE_WHT_TCC)
        ttk.Checkbutton(wf, text="Generate missing TCC numbers", variable=self.gen_tcc).pack(
            anchor="w", pady=(8, 0))
        ttk.Label(wf, text="creates them on the client's URA account:\n"
                           "ticks rows with no TCC number, presses\n"
                           "Generate TCC, then reads the new numbers",
                  style="Muted.TLabel", justify="left").pack(anchor="w", padx=(22, 0))

        pf = ttk.LabelFrame(f, text=" PAYE (Pay As You Earn) ", padding=8)
        pf.grid(row=2, column=1, sticky="nsew", padx=(0, 8))
        self.paye = tk.BooleanVar(value=True)
        self.paye_gen = tk.BooleanVar(value=self.DEFAULT_GENERATE_PAYE_TCC)
        self.paye_mw = tk.BooleanVar(value=True)
        ttk.Checkbutton(pf, text="PAYE tax credit certificates", variable=self.paye).pack(anchor="w")
        ttk.Checkbutton(pf, text="generate missing TCC numbers",
                        variable=self.paye_gen).pack(
            anchor="w", padx=(22, 0))
        ttk.Label(pf, text="creates them on URA", style="Muted.TLabel").pack(anchor="w", padx=(44, 0))
        ttk.Checkbutton(pf, text="each employee's month-wise details", variable=self.paye_mw).pack(
            anchor="w", padx=(22, 0))
        self.fy_note = ttk.Label(pf, text="", style="Muted.TLabel", justify="left", wraplength=280)
        self.fy_note.pack(anchor="w", padx=(22, 0), pady=(6, 0))

        of = ttk.LabelFrame(f, text=" Saved files ", padding=8)
        of.grid(row=2, column=2, sticky="nsew")
        self.opts = {}
        for key, label, default in [
                ("save_excel", "Excel workbook", True),
                ("save_html", "Portal page copies (HTML)", True),
                ("save_csv", "One flat CSV (WHT)", False),
                ("save_month_excels", "A workbook per month (WHT)", False),
                ("continue_on_error", "Carry on if a month fails", True)]:
            v = tk.BooleanVar(value=default)
            self.opts[key] = v
            ttk.Checkbutton(of, text=label, variable=v).pack(anchor="w")
        ttk.Label(of, text="No PDFs are made for WHT or PAYE.", style="Muted.TLabel").pack(
            anchor="w", pady=(6, 0))
        return f

    def on_period_change(self, pf, pt):
        from engines.paye_engine import financial_years, fy_label
        try:
            ys = financial_years(pf, pt)
            self.fy_note.configure(text="PAYE searches:\n" + "\n".join(fy_label(y) for y in ys))
        except Exception:
            self.fy_note.configure(text="")

    def collect(self):
        if not self.shared_period:
            raise ValueError("Certificates: set the period first.")
        pf, pt = self.shared_period
        a = datetime.datetime.strptime(pf, "%Y-%m")
        b = datetime.datetime.strptime(pt, "%Y-%m")
        kinds = [k for k, v in self.gen_vars.items() if v.get()]
        if not kinds and not self.paye.get():
            raise ValueError("Certificates: tick Payee, Withholding Agent or PAYE.")
        return {"from": (a.year, a.month), "to": (b.year, b.month), "period": (pf, pt),
                "generate_for": kinds, "generate_tcc": self.gen_tcc.get(),
                "paye": self.paye.get(), "paye_generate": self.paye_gen.get(),
                "paye_monthwise": self.paye_mw.get(),
                "save_table_pdf": False, "save_page_pdf": False,
                **{k: v.get() for k, v in self.opts.items()}}

    CODES = {"Payee": "D", "Payer": "W"}     # live portal values (3 Oct 2026): D = Payee, W = Withholding Agent

    def run(self, ctx, client, options, page=None):
        from engines import wht_engine as we
        if page is None:
            raise RuntimeError("Certificates need a logged-in session (core.session.open_session).")
        kinds = options["generate_for"]
        if isinstance(kinds, str):          # settings saved by an older version
            kinds = [kinds]
        total = {"files": 0, "months_ok": 0, "months_error": 0}
        failed_kinds, paye_failed = [], []
        if kinds:
            bot = we.WhtBot(page, ctx.log, ctx.check, screenshot_dir=str(ctx.log_path.parent))
            bot.dismiss_popups()
            bot.open_wht_page()
            for i, kind in enumerate(kinds):
                if i > 0:
                    bot.open_wht_page(quiet=True)
                res = self._run_one(ctx, client, options, bot, kind)
                for k in total:
                    total[k] += res.get(k, 0)
                if res.get("months_error"):
                    failed_kinds.append(kind)
        if options.get("paye"):
            from core.session import go_home
            from engines import paye_engine
            if kinds:
                go_home(ctx, page, client, before="PAYE certificates")
            out = ctx.output_root / "Certificates" / client["login"] / "PAYE"
            pf, pt = options["period"]
            files = paye_engine.run(
                page, client["login"], pf, pt, out, ctx.log, ctx.check,
                lambda item, status, detail="", file="": ctx.record(item, status, detail, file),
                generate=options.get("paye_generate", self.DEFAULT_GENERATE_PAYE_TCC),
                monthwise=options.get("paye_monthwise", True), debug_dir=ctx.log_path.parent,
                reopen=lambda: go_home(ctx, page, client, before="searching PAYE certificates again"),
                failed=paye_failed, years=options.get("_paye_years"))
            total["files"] += len(files)
        if (failed_kinds or paye_failed) and not ctx.final_try:     # 10.15: once more at the end of the client
            from engines import paye_engine
            what = ", ".join([f"WHT {k}" for k in failed_kinds]
                             + [f"PAYE {paye_engine.fy_label(y)}" for y in paye_failed])
            ctx.log(f"Certificates not finished ({what}) - tried again after this client's other steps.")
            ctx.after_client.append((f"Certificates - second try for {what}",
                                     lambda pg: self.second_try(ctx, client, options, pg, failed_kinds,
                                                                paye_failed)))
        return total

    def second_try(self, ctx, client, options, page, kinds, paye_years):
        """WHT: the whole type again (its workbook holds every month); PAYE: only the failed years.
        The first try's items are updated ('data captured on second try'), not listed twice."""
        return self.run(ctx, client, dict(options, generate_for=list(kinds), paye=bool(paye_years),
                                          _paye_years=list(paye_years)), page)

    def _run_one(self, ctx, client, options, bot, kind):
        from engines import wht_engine as we
        from engines import wht_outputs as wo
        tin = client["login"]
        out_dir = ctx.output_root / "Certificates" / tin / "WHT"
        out_dir.mkdir(parents=True, exist_ok=True)
        (y1, m1), (y2, m2) = options["from"], options["to"]
        months = [wo.MonthResult(y, m) for y, m in wo.month_sequence(y1, m1, y2, m2)]
        gen_code = self.CODES[kind]
        label = "Payee (Deductee)" if kind == "Payee" else "Withholding Agent (Payer)"
        ctx.log(f"WHT Receivable - {label}: {len(months)} month(s), "
                f"{months[0].title} to {months[-1].title}.")
        files = []

        def on_result(r, i):
            k = {"ok": "ok", "empty": "", "error": "err", "skipped": "warn"}.get(r.status, "")
            ctx.log(f"  {kind} {r.title}: {r.status} - {r.rows} row(s) {r.message}", k)
        stopped = False
        try:
            we.run_months(bot, months, want_pdf=options.get("save_page_pdf", False), want_screenshot=False,
                          want_html=options["save_html"], generate_for=gen_code,
                          continue_on_error=options["continue_on_error"], on_result=on_result,
                          generate_tcc=options.get("generate_tcc", self.DEFAULT_GENERATE_WHT_TCC))
        except (we.Cancelled, Stopped):
            stopped = True                  # still save the months already captured
            ctx.log("Stopped - saving the WHT months captured so far.", "warn")
        if stopped:
            for r in months:                # make the gap visible in every output
                if r.status == "pending":
                    r.status, r.message = "skipped", "Not run - stopped by you"
            if not any(r.status in ("ok", "empty", "error") for r in months):
                raise Stopped()

        # ---- files: one set per type, the type in every file name ----
        stamp = f"{y1}{m1:02d}-{y2}{m2:02d}"
        prefix = f"WHT_{tin}_{kind}_{stamp}"
        with_data = [r for r in months if r.df is not None and len(r.df)]
        if options["save_excel"]:
            files.append(wo.save_excel(months, str(out_dir / f"{prefix}.xlsx"), tin=tin,
                                       gen_label=label))
            if kind == "Payee":                   # 10.13: certificates against the Income Tax ledger
                self._check_vs_ledger(ctx, tin, months, files[-1])
        if options.get("save_table_pdf") and with_data:
            files.append(wo.save_table_pdf(months, str(out_dir / f"{prefix}_table.pdf"), tin=tin,
                                           gen_label=label))
        if options.get("save_page_pdf"):
            merged = wo.merge_pdfs(months, str(out_dir / f"{prefix}_pages.pdf"))
            if merged:
                files.append(merged)
        if options["save_html"]:
            for r in months:
                if r.html:
                    pth = out_dir / "pages" / f"{prefix}_{r.label}.html"
                    pth.parent.mkdir(exist_ok=True)
                    pth.write_text(r.html, encoding="utf-8")
                    files.append(str(pth))
        if options["save_csv"]:
            c = wo.save_csv(months, str(out_dir / f"{prefix}.csv"))
            if c:
                files.append(c)
        if options["save_month_excels"]:
            files += wo.save_month_excels(months, str(out_dir / "months"), prefix=f"WHT_{tin}_{kind}")
        for r in months:
            st = {"ok": "data captured", "empty": "no records", "error": "failed",
                  "skipped": "skipped"}.get(r.status, r.status)
            ctx.record(f"WHT {kind} {r.title}", st, f"{r.rows} row(s). {r.message}".strip())
            if "still without a TCC number" in (r.message or ""):
                ctx.record(f"WHT {kind} {r.title} - TCC numbers", "incomplete", r.message)
        for f in files:
            ctx.log(f"Saved: {f}", "ok")
            ctx.record(os.path.basename(f), "saved", label, f)
        errors = sum(1 for r in months if r.status == "error")
        if stopped:
            raise Stopped()
        return {"files": len(files), "months_ok": len(with_data), "months_error": errors}

    @staticmethod
    def _check_vs_ledger(ctx, tin, months, workbook):
        """Sheet 'Check vs ledger' in the Payee workbook; a certificate missing from the Income Tax ledger is a
        'check' item on the Review page.  Never stops the step."""
        from engines import wht_crosscheck as wc
        try:
            result = wc.crosscheck(months, ctx.output_root / "Ledgers" / tin)
            wc.add_sheet(workbook, result)
        except Exception as e:
            ctx.log(f"  Certificates not checked against the Income Tax ledger ({type(e).__name__}).", "warn")
            return
        if result is None:
            ctx.log("  Certificates not checked against the Income Tax ledger: no Income Tax Level 3 file "
                    "(tick Income Tax and its Level 3 on the Ledgers page).")
            return
        certs, left = result[0], result[1]
        missing = certs[certs["Found in Income Tax ledger"] == "No"] if len(certs) else certs
        ctx.log(f"  Checked against the Income Tax ledger: {len(certs) - len(missing)} of {len(certs)} "
                f"certificate(s) found; {len(left)} ledger WHT credit line(s) without a certificate "
                "(sheet 'Check vs ledger').", "warn" if len(missing) or left else "ok")
        tcc_col = next((c for c in certs.columns if "tcc" in str(c).lower()), None)
        amt_col = next((c for c in certs.columns if "withheld" in str(c).lower()), None)
        for _, c in missing.iterrows():
            what = f"TCC {c.get(tcc_col, '')}".strip() if tcc_col and str(c.get(tcc_col, "")).strip() else "no TCC No."
            ctx.record(f"WHT Payee {c['Month']} - {what}", "check",
                       f"certificate not found in the Income Tax ledger (Amount Withheld "
                       f"{c.get(amt_col, '') if amt_col else ''})", workbook)
        for ln in left:
            ctx.record(f"Income Tax ledger WHT credit - {ln['file']}", "check",
                       f"no WHT certificate matches: {ln['text'][:150]}", workbook)
