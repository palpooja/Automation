"""
URA Suite - one window for all the URA e-Tax downloaders.

Left: a vertical menu.  Each module has a tick box (include it in
"Run selected") and a name (click to open its options page).
Right: the open module's page, the shared client list, the run buttons,
the yellow question bar and the activity log.

Run order: client by client.  Each client is logged in ONCE
(core.session.open_session); every ticked module then runs on that same
browser page in menu order; then logout.  One security code per client.
"""
import datetime
import os
import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

VERSION = "10.18"            # shown in the window title and at the top of every run log

sys.path.insert(0, str(Path(__file__).resolve().parent))

from core import store                                    # noqa: E402
from core.runner import BrowserClosed, RunContext, Stopped, browser_gone, needs_review  # noqa: E402
from modules.returns import ReturnsModule                 # noqa: E402
from modules.ledger import LedgerModule                   # noqa: E402
from modules.wht_receivable import WhtReceivableModule    # noqa: E402
from modules.settings_page import SettingsPage, suite_settings  # noqa: E402
from core import theme                                    # noqa: E402
from core import voice                                    # noqa: E402  spoken 'Job completed'
from core.widgets import WebButton, TextField, PasswordField  # noqa: E402

MODULES = [ReturnsModule(), LedgerModule(), WhtReceivableModule()]

NAVY, NAVY_DARK, PAPER, CARD, LINE = "#1F4E79", "#173A5A", "#F3F5F8", "#FFFFFF", "#DDE3EA"
TEXT, MUTED, GREEN, RED, AMBER = "#1B2733", "#66737F", "#1E8E3E", "#C62828", "#C77700"
FONT = ("Segoe UI", 10)


QUEUE_ROWS = 12          # rows in the left 'This run' list; more clients are counted, not listed


class App(tk.Tk):
    def __init__(self):
        from core import display
        display.make_dpi_aware()               # sharp text and correct sizes under Windows scaling
        theme.register_fonts()                 # bundled typefaces, this program only
        super().__init__()
        self.settings = store.load_settings()
        name = self.settings.get("suite", {}).get("theme", theme.DEFAULT_THEME)
        self.T = T = theme.resolve(name, self)
        theme.CURRENT = T
        global FONT, TEXT, MUTED, GREEN
        FONT, TEXT, MUTED, GREEN = T["FONT"], T["ink"], T["muted"], T["accent"]
        self.title(f"{store.APP_NAME} {VERSION} - Returns, Ledgers, WHT")
        # fits any screen: 1240x820 where there is room, never larger than the visible screen
        self.S = display.place(self, 1240, 820, 900, 560, maximize_if_small=True)
        self.configure(bg=T["page_bg"])
        try:
            self.iconbitmap(str(store.resource("ura_icon.ico")))
        except Exception:
            pass
        store.ensure_dirs()
        self.clients = store.load_clients()
        self._load_warnings = list(store.LOAD_WARNINGS)
        try:                                   # housekeeping at start-up: no run time is spent
            self._pruned = store.prune_logs(int(self.settings.get("suite", {}).get("log_keep_days", 30)))
        except (OSError, ValueError):
            self._pruned = 0
        try:                                   # ledger files saved before 10.2 that Excel refused
            from engines.excel_files import repair_folder
            self._repaired = repair_folder(Path(self.settings["output_root"]) / "Ledgers")
        except OSError:
            self._repaired = 0
        self.ticked_logins = {c["login"] for c in self.clients}
        self.ui_q, self.reply_q = queue.Queue(), queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None
        self.current = None
        self.client_states = {}
        self._style()
        self._build()
        self._pump()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------ look
    def _style(self):
        T = self.T
        s = ttk.Style(self)
        try:
            s.theme_use("vista")
        except Exception:
            pass
        s.configure(".", font=T["FONT"], background=T["card_bg"], foreground=T["ink"])
        for st in ("TFrame", "TLabel", "TCheckbutton", "TLabelframe", "TRadiobutton"):
            s.configure(st, background=T["card_bg"], foreground=T["ink"])
        s.configure("TLabelframe", bordercolor=T["line"])
        s.configure("TLabelframe.Label", background=T["card_bg"], foreground=T["frame_label"],
                    font=T["FONT_SUB"])
        s.configure("Head.TLabel", font=T["FONT_HEAD"], foreground=T["head"])
        s.configure("Muted.TLabel", foreground=T["muted"])
        s.map("TCheckbutton", background=[("disabled", T["card_bg"]), ("active", T["card_bg"])],
              foreground=[("disabled", T["muted"])])
        s.configure("Treeview", font=T["FONT"], rowheight=int(24 * self.S), background=T["card_bg"],
                    fieldbackground=T["card_bg"], foreground=T["ink"], bordercolor=T["line"])
        s.configure("Treeview.Heading", font=T["FONT_BOLD"])

    def _button(self, parent, text, cmd, primary=False, kind=None):
        return WebButton(parent, text=text, command=cmd,
                         kind=kind or ("primary" if primary else "secondary"),
                         surface=self.T["card_bg"], width=200)

    # ------------------------------------------------------------ layout
    def _build(self):
        T = self.T
        rail_bg = T["rail_bg"]
        # ---- left: vertical menu ----
        menu = tk.Frame(self, bg=rail_bg, width=int(236 * self.S))
        menu.pack(side="left", fill="y")
        menu.pack_propagate(False)
        tk.Label(menu, text="URA Suite", bg=rail_bg, fg=T["rail_fg"], font=T["FONT_BRAND"],
                 anchor="w", padx=16, pady=16).pack(fill="x")
        tk.Label(menu, text="Steps for each client", bg=rail_bg, fg=T["rail_muted"], font=T["FONT_SMALL"],
                 anchor="w", padx=16).pack(fill="x", pady=(0, 2))
        self.menu_vars, self.menu_btns, self.menu_rows = {}, {}, {}
        for step_no, m in enumerate(MODULES, 1):
            row = tk.Frame(menu, bg=rail_bg)
            row.pack(fill="x")
            v = tk.BooleanVar(value=self.settings["modules"].get(m.id, {}).get("ticked", True))
            self.menu_vars[m.id] = v
            cb = tk.Checkbutton(row, variable=v, bg=rail_bg, activebackground=rail_bg,
                                selectcolor=T["rail_active_bg"] if T["rail_dark"] else "#FFFFFF",
                                fg=T["rail_fg"], highlightthickness=0, bd=0, relief="flat")
            cb.pack(side="left", padx=(12, 0))
            b = tk.Button(row, text=f"{step_no}   {m.title}", anchor="w", bg=rail_bg, fg=T["rail_fg"], bd=0,
                          relief="flat", highlightthickness=0, activebackground=T["rail_active_bg"],
                          activeforeground=T["rail_active_fg"], font=T["FONT_RAIL"],
                          padx=8, pady=9, cursor="hand2", command=lambda mod=m: self.show(mod))
            b.pack(side="left", fill="x", expand=True)
            self.menu_btns[m.id] = b
            for w in (b,):
                w.bind("<Enter>", lambda e, mid=m.id: self._rail_hover(mid, True))
                w.bind("<Leave>", lambda e, mid=m.id: self._rail_hover(mid, False))
            self.menu_rows[m.id] = (row, cb)
        rrow = tk.Frame(menu, bg=rail_bg)
        rrow.pack(fill="x")
        tk.Label(rrow, text="", bg=rail_bg, width=2).pack(side="left", padx=(12, 0))
        rb = tk.Button(rrow, text=f"{len(MODULES) + 1}   Review", anchor="w", bg=rail_bg, fg=T["rail_fg"],
                       bd=0, relief="flat", highlightthickness=0, activebackground=T["rail_active_bg"],
                       activeforeground=T["rail_active_fg"], font=T["FONT_RAIL"], padx=8, pady=9,
                       cursor="hand2", command=lambda: self.show(self.review))
        rb.pack(side="left", fill="x", expand=True)
        self.menu_btns["review"] = rb
        self.menu_rows["review"] = (rrow, rrow.winfo_children()[0])
        tk.Frame(menu, bg=T["rail_rule"], height=1).pack(fill="x", pady=12, padx=14)
        for text, cmd in (("Clients", self.manage_clients), ("History", self.show_history),
                          ("Settings", lambda: self.show(self.settings_page)),
                          ("Open portal (manual)", self.open_manual),
                          ("Output folder", self.choose_output), ("Open output folder", self.open_output)):
            tk.Button(menu, text=text, anchor="w", bg=rail_bg, fg=T["rail_muted"], bd=0,
                      relief="flat", highlightthickness=0, activebackground=T["rail_active_bg"],
                      activeforeground=T["rail_active_fg"], font=T["FONT"], padx=18, pady=5,
                      cursor="hand2", command=cmd).pack(fill="x")
        self.queue_frame = None
        if T["queue_in_rail"]:
            tk.Frame(menu, bg=T["rail_rule"], height=1).pack(fill="x", pady=12, padx=14)
            tk.Label(menu, text="This run", bg=rail_bg, fg=T["rail_fg"], font=T["FONT_SUB"],
                     anchor="w", padx=16).pack(fill="x")
            self.queue_frame = tk.Frame(menu, bg=rail_bg)
            self.queue_frame.pack(fill="x", padx=16, pady=(4, 0))
        self.menu_status = tk.Label(menu, text="", bg=rail_bg, fg=T["rail_muted"], wraplength=int(200 * self.S),
                                    justify="left", anchor="sw", padx=16, pady=12, font=T["FONT"])
        self.menu_status.pack(side="bottom", fill="x")

        # ---- right ----
        from core.display import ScrollArea
        self.right_scroll = ScrollArea(self, bg=T["page_bg"])      # scrolls only on a short screen
        self.right_scroll.pack(side="left", fill="both", expand=True)
        right = self.right_scroll.inner
        edge = 0 if T["tab_join"] else 1
        self.page_host = tk.Frame(right, bg=T["card_bg"], highlightbackground=T["line"],
                                  highlightthickness=edge)
        self.page_host.pack(fill="x", padx=(0 if T["tab_join"] else 12, 12), pady=(12, 6))
        self.pages = {}
        for m in MODULES:
            p = m.build_panel(self.page_host)
            self.pages[m.id] = p
        self.settings_page = SettingsPage(self.settings, None)
        self.pages["settings"] = self.settings_page.build_panel(self.page_host)
        self.review = ReviewPage(self)
        self.pages["review"] = self.review.build_panel(self.page_host)

        # clients + run buttons
        cf = tk.Frame(right, bg=T["card_bg"], highlightbackground=T["line"], highlightthickness=1)
        cf.pack(fill="x", padx=12, pady=6)
        left = tk.Frame(cf, bg=T["card_bg"])
        left.pack(side="left", fill="both", expand=True, padx=10, pady=8)
        prow = tk.Frame(left, bg=T["card_bg"])
        prow.pack(fill="x", pady=(0, 6))
        tk.Label(prow, text="Period", bg=T["card_bg"], fg=T["head"], font=T["FONT_SUB"]).pack(side="left")
        saved = self.settings.get("period") or {}
        d_from, d_to = default_period()
        self.period_from = tk.StringVar(value=saved.get("from", d_from))
        self.period_to = tk.StringVar(value=saved.get("to", d_to))
        for lbl, var in (("from", self.period_from), ("to", self.period_to)):
            tk.Label(prow, text=lbl, bg=T["card_bg"], fg=T["ink"], font=T["FONT"]).pack(side="left", padx=(12, 4))
            ttk.Entry(prow, textvariable=var, width=9).pack(side="left")
        tk.Label(prow, text="year-month", bg=T["card_bg"], fg=T["muted"], font=T["FONT_SMALL"]).pack(
            side="left", padx=(6, 10))
        WebButton(prow, text="This financial year", command=self._fy_now, surface=T["card_bg"],
                  padx=10, pady=3).pack(side="left", padx=2)
        WebButton(prow, text="Last 18 months", command=self._last18, surface=T["card_bg"],
                  padx=10, pady=3).pack(side="left", padx=2)
        self.period_msg = tk.Label(prow, text="", bg=T["card_bg"], fg=T["danger"], font=T["FONT_SMALL"])
        self.period_msg.pack(side="left", padx=8)
        for var in (self.period_from, self.period_to):
            var.trace_add("write", lambda *_: self._period_changed())
        tk.Label(left, text="Clients - tick the ones to include", bg=T["card_bg"], fg=T["head"],
                 font=T["FONT_SUB"]).pack(anchor="w", pady=(0, 4))
        self.tree = ttk.Treeview(left, columns=("tick", "name", "login"), show="headings",
                                 height=3, selectmode="browse")
        for col, head, w in (("tick", "", 34), ("name", "Client", 260), ("login", "Login ID (TIN)", 160)):
            self.tree.heading(col, text=head)
            self.tree.column(col, width=int(w * self.S), anchor="center" if col == "tick" else "w",
                             stretch=(col == "name"))
        self.tree.pack(fill="x")
        self.tree.bind("<Button-1>", self._tree_click)
        self.count_lbl = tk.Label(left, text="", bg=T["card_bg"], fg=T["muted"], font=T["FONT_SMALL"])
        self.count_lbl.pack(anchor="w", pady=(4, 0))
        self._refresh_clients()

        btns = tk.Frame(cf, bg=T["card_bg"])
        btns.pack(side="right", padx=12, pady=10, anchor="n")
        self._clients_row = (cf, left, btns)
        self._stacked = False
        self.right_scroll.canvas.bind("<Configure>", lambda e: self._reflow(e.width), add="+")
        self.run_sel_btn = self._button(btns, "Run selected modules",
                                        lambda: self.start(all_ticked=True), primary=True)
        self.run_sel_btn.pack(fill="x", pady=(0, 2))
        self.run_one_btn = self._button(btns, "Run this page only", lambda: self.start(all_ticked=False))
        self.run_one_btn.pack(fill="x", pady=(0, 2))
        self.stop_btn = self._button(btns, "Stop", self.stop, kind="danger")
        self.stop_btn.configure(state="disabled")
        self.stop_btn.pack(fill="x")

        # question bar (hidden until needed) - the one loud element
        self.ask = tk.Frame(right, bg=T["prompt_bg"])
        tk.Frame(self.ask, bg=T["prompt_line"], width=6).pack(side="left", fill="y")
        inner = tk.Frame(self.ask, bg=T["prompt_bg"])
        inner.pack(side="left", fill="both", expand=True, padx=14, pady=10)
        self.ask_who = tk.Label(inner, text="", bg=T["prompt_bg"], fg=T["prompt_fg"],
                                font=T["FONT_SUB"], anchor="w")
        self.ask_who.pack(anchor="w")
        self.ask_msg = tk.Label(inner, text="", bg=T["prompt_bg"], fg=T["prompt_fg"],
                                font=T["FONT"], wraplength=int(860 * self.S), justify="left")
        self.ask.bind("<Configure>", lambda e: self.ask_msg.configure(wraplength=max(300, e.width - 60)))
        self.ask_msg.pack(anchor="w", pady=(2, 8))
        self.ask_row = tk.Frame(inner, bg=T["prompt_bg"])
        self.ask_row.pack(anchor="w")

        # log
        lf = tk.Frame(right, bg=T["page_bg"])
        lf.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self.log_frame = lf
        head = tk.Frame(lf, bg=T["page_bg"])
        head.pack(fill="x", pady=(0, 4))
        tk.Label(head, text="Activity", bg=T["page_bg"], fg=T["head"],
                 font=T["FONT_SUB"]).pack(side="left")
        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(head, textvariable=self.status_var, bg=T["page_bg"], fg=T["muted"],
                 font=T["FONT"]).pack(side="right")
        self.logbox = ScrolledText(lf, height=8, font=T["FONT_LOG"], state="disabled",   # grows to fill
                                   bg=T["log_bg"], fg=T["log_fg"], relief="flat",
                                   highlightthickness=1, highlightbackground=T["line"],
                                   padx=8, pady=6)
        self.logbox.pack(fill="both", expand=True)
        for tag, key in (("ok", "log_ok"), ("err", "log_err"), ("warn", "log_warn")):
            self.logbox.tag_configure(tag, foreground=T[key])

        self._period_changed()
        self.show(MODULES[0])
        self.log(f"Clients are stored with {store.credential_store_kind()} in {store.DATA_DIR}.")
        self.log(f"Output folder: {self.settings['output_root']}")
        if self._repaired:
            self.log(f"Repaired {self._repaired} ledger Excel file(s) that Excel could not open "
                     "(saved by an earlier version). The portal's originals are in 'portal_original'.", "ok")
        if self._pruned:
            self.log(f"Deleted {self._pruned} old screenshot/page-copy/log file(s) from the log folder.")
        for w in self._load_warnings:
            self.log(w, "warn")
        if self._load_warnings:
            self.after(400, lambda: messagebox.showwarning("Clients", "\n\n".join(self._load_warnings)))

    def _reflow(self, width):
        """Narrow window: the Run / Stop buttons go under the client list instead of beside it,
        so they never end up past the right edge (1024-px screens)."""
        cf, left, btns = self._clients_row
        need = left.winfo_reqwidth() + btns.winfo_reqwidth() + int(60 * self.S)
        stack = width < need if not self._stacked else width < need + int(40 * self.S)
        if stack == self._stacked:
            return
        self._stacked = stack
        left.pack_forget(); btns.pack_forget()
        buttons = (self.run_sel_btn, self.run_one_btn, self.stop_btn)
        for b in buttons:
            b.pack_forget()
        if stack:
            left.pack(side="top", fill="x", padx=10, pady=(8, 0))
            btns.pack(side="top", anchor="w", padx=10, pady=(6, 10))
            for b in buttons:
                b.pack(side="left", padx=(0, 6))
        else:
            left.pack(side="left", fill="both", expand=True, padx=10, pady=8)
            btns.pack(side="right", padx=12, pady=10, anchor="n")
            for b in buttons[:2]:
                b.pack(fill="x", pady=(0, 2))
            self.stop_btn.pack(fill="x")

    # ------------------------------------------------------------ run queue (design C)
    def _render_queue(self):
        if not self.queue_frame:
            return
        T = self.T
        for w in self.queue_frame.winfo_children():
            w.destroy()
        rows = [c for c in self.clients if c["login"] in self.ticked_logins]
        if not rows:
            tk.Label(self.queue_frame, text="Tick clients to build the run.", bg=T["rail_bg"],
                     fg=T["rail_muted"], font=T["FONT_SMALL"], anchor="w").pack(fill="x")
        if len(rows) > QUEUE_ROWS:              # many clients: counts + the ones that matter now
            states = [self.client_states.get(c["login"], "waiting") for c in rows]
            counts = {s: states.count(s) for s in dict.fromkeys(states)}
            tk.Label(self.queue_frame, text=f"{len(rows)} clients - " + ", ".join(f"{n} {s}" for s, n in counts.items()),
                     bg=T["rail_bg"], fg=T["rail_fg"], font=T["FONT_SMALL"], anchor="w", justify="left",
                     wraplength=int(200 * self.S)).pack(fill="x", pady=(0, 4))
            order = ([c for c, s in zip(rows, states) if s not in ("waiting", "done", "failed", "stopped")] +
                     [c for c, s in zip(rows, states) if s == "waiting"] +
                     [c for c, s in zip(rows, states) if s in ("done", "failed", "stopped")][::-1])
            rows = order[:QUEUE_ROWS - 2]
        for c in rows:
            st = self.client_states.get(c["login"], "waiting")
            r = tk.Frame(self.queue_frame, bg=T["rail_bg"])
            r.pack(fill="x", pady=2)
            col = T["state_colours"].get(st, T["rail_muted"])
            tk.Label(r, text="\u25cf", bg=T["rail_bg"], fg=col, font=T["FONT"]).pack(side="left")
            tk.Label(r, text=c["name"], bg=T["rail_bg"], fg=T["rail_fg"], font=T["FONT"],
                     anchor="w").pack(side="left", padx=(4, 0))
            tk.Label(r, text=st, bg=T["rail_bg"], fg=col, font=T["FONT_SMALL"]).pack(side="right")

    def _set_client_state(self, login, state):
        self.client_states[login] = state
        self._render_queue()

    def _fy_now(self):
        a, b = default_period()
        self.period_from.set(a); self.period_to.set(b)

    def _last18(self):
        today = datetime.date.today()
        y, m = today.year, today.month - 17
        while m < 1:
            y, m = y - 1, m + 12
        self.period_from.set(f"{y}-{m:02d}"); self.period_to.set(f"{today.year}-{today.month:02d}")

    def _period_changed(self):
        """Validate the shared period and hand it to every step."""
        pf, pt = self.period_from.get().strip(), self.period_to.get().strip()
        try:
            a = datetime.datetime.strptime(pf, "%Y-%m")
            b = datetime.datetime.strptime(pt, "%Y-%m")
            if a > b:
                raise ValueError("'from' is after 'to'")
        except ValueError as e:
            self.period_msg.configure(text="type year-month, e.g. 2025-07" if "does not match" in str(e)
                                      else str(e))
            for m in MODULES:
                m.shared_period = None
            return False
        self.period_msg.configure(text="")
        for m in MODULES:
            m.shared_period = (pf, pt)
            if hasattr(m, "on_period_change"):
                try:
                    m.on_period_change(pf, pt)
                except Exception:
                    pass
        self.settings["period"] = {"from": pf, "to": pt}
        try:
            store.save_settings(self.settings)
        except Exception:
            pass
        return True

    def _rail_hover(self, mid, on):
        if self.current is not None and getattr(self.current, "id", None) == mid:
            return
        T = self.T
        from core.widgets import _shade
        bg = (_shade(T["rail_bg"], 1.08) if T["rail_dark"] else _shade(T["rail_bg"], 0.96)) \
            if on else T["rail_bg"]
        b = self.menu_btns[mid]
        row, cb = self.menu_rows[mid]
        for w in (b, row, cb):
            w.configure(bg=bg)
        cb.configure(activebackground=bg)

    # ------------------------------------------------------------ pages
    def show(self, module):
        T = self.T
        for mid, p in self.pages.items():
            p.pack_forget()
        self.pages[module.id].pack(fill="both", expand=True)
        for mid, b in self.menu_btns.items():
            on = mid == module.id
            bg = T["rail_active_bg"] if on else T["rail_bg"]
            b.configure(bg=bg, fg=T["rail_active_fg"] if on else T["rail_fg"])
            row, cb = self.menu_rows[mid]
            row.configure(bg=bg)
            cb.configure(bg=bg, activebackground=bg)
        if module.id in ("settings", "review"):
            self.run_one_btn.configure(text="Run this page only", state="disabled")
            return
        self.current = module
        if not (self.worker and self.worker.is_alive()):
            self.run_one_btn.configure(state="normal")
        self.run_one_btn.configure(text=f"Run '{module.title}' only")

    # ------------------------------------------------------------ clients
    def _refresh_clients(self):
        self.tree.delete(*self.tree.get_children())
        for c in self.clients:
            tick = "\u2611" if c["login"] in self.ticked_logins else "\u2610"
            self.tree.insert("", "end", values=(tick, c["name"], c["login"]))
        n = sum(1 for c in self.clients if c["login"] in self.ticked_logins)
        self.count_lbl.configure(text=f"{n} of {len(self.clients)} ticked")
        if getattr(self, "queue_frame", None) is not None:
            self._render_queue()

    def _tree_click(self, event):
        if self.tree.identify_column(event.x) != "#1":
            return
        row = self.tree.identify_row(event.y)
        if not row:
            return
        login = self.tree.item(row, "values")[2]
        self.ticked_logins ^= {login}
        self._refresh_clients()

    def manage_clients(self):
        ClientsDialog(self)

    # ------------------------------------------------------------ settings
    def choose_output(self):
        d = filedialog.askdirectory(initialdir=self.settings["output_root"],
                                    title="Folder where all downloads are saved")
        if d:
            self.settings["output_root"] = d
            store.save_settings(self.settings)
            self.log(f"Output folder: {d}")

    def open_output(self):
        p = self.settings["output_root"]
        os.makedirs(p, exist_ok=True)
        if sys.platform.startswith("win"):
            os.startfile(p)
        else:
            subprocess.Popen(["xdg-open", p])

    def show_history(self):
        HistoryDialog(self)

    # ------------------------------------------------------------ running
    def start(self, all_ticked):
        if self.worker and self.worker.is_alive():
            return
        mods = ([m for m in MODULES if self.menu_vars[m.id].get()] if all_ticked
                else [self.current])
        if not mods:
            messagebox.showwarning("Nothing ticked", "Tick at least one module in the menu.")
            return
        run_list = [c for c in self.clients if c["login"] in self.ticked_logins]
        if not run_list:
            messagebox.showwarning("No clients", "Tick at least one client (Clients to add).")
            return
        try:
            if not self._period_changed():
                raise ValueError("Set the period (year-month, e.g. 2025-07 to 2026-06).")
            options = {m.id: m.collect() for m in mods}
        except ValueError as e:
            messagebox.showerror("Check the options", str(e))
            return
        for m in mods:
            self.settings["modules"].setdefault(m.id, {})["ticked"] = self.menu_vars[m.id].get()
        store.save_settings(self.settings)

        no_pw = [c for c in run_list if not c.get("password")]
        if no_pw:
            messagebox.showerror("Passwords needed", "Type the password again (Clients > Edit) for:\n\n"
                                 + "\n".join(f"  - {c['name']} ({c['login']})" for c in no_pw)
                                 + "\n\nThe saved one cannot be decrypted on this Windows account.")
            return
        su = suite_settings(self.settings)
        out_root = Path(self.settings["output_root"])
        problem = self._preflight(out_root, su)
        if problem:
            messagebox.showerror("Cannot start", problem)
            return
        self._apply_engine_settings(su)

        # resume: drop client+module pairs already completed today
        plan = [(c, m) for c in run_list for m in mods]
        skipped = []
        if su["skip_done_today"]:
            # Only a CLEAN completion counts: status exactly 'Completed' (not
            # 'Completed - N item(s) failed') AND at least one file saved.
            today = datetime.date.today().strftime("%d/%m/%Y")
            done = {}
            for h in store.load_history():
                if (h.get("status") == "Completed" and str(h.get("when", "")).startswith(today)
                        and int(h.get("files") or 0) > 0):
                    done[(h.get("login"), h.get("module"))] = h
            cand = [(c, m) for c, m in plan if (c["login"], m.title) in done]
            if cand:
                lines = "\n".join(f"  - {c['name']}: {m.title} "
                                   f"({done[(c['login'], m.title)].get('files')} file(s) at "
                                   f"{str(done[(c['login'], m.title)].get('when'))[-5:]})"
                                   for c, m in cand)
                ans = messagebox.askyesnocancel(
                    "Already completed today",
                    "These were already completed today without failures:\n\n" + lines +
                    "\n\nYes = skip them (faster)\nNo = run them again\nCancel = do not start")
                if ans is None:
                    return
                if ans:
                    skipped = cand
                    plan = [(c, m) for c, m in plan if (c, m) not in cand]
            if skipped and not plan:
                messagebox.showinfo("Nothing to do", "Everything ticked was already completed today.")
                return

        self.stop_event.clear()
        out_root.mkdir(parents=True, exist_ok=True)
        stamp = f"{datetime.datetime.now():%Y%m%d_%H%M%S}"
        log_path = store.LOG_DIR / f"run_{stamp}.txt"
        ctx = RunContext(self.ui_q, self.reply_q, self.stop_event, out_root, log_path)
        ctx.voice_choice = su.get("voice", voice.DEFAULT)
        ctx.login_timeout_s = su["login_timeout_s"]
        from core import net
        from engines import ledger_engine as le, wht_engine as we, returns_pw as rp
        le.ONLINE_WAIT = we.ONLINE_WAIT = rp.ONLINE_WAIT = lambda: net.wait_online(ctx.log, ctx.check)
        self.client_states = {}
        self._render_queue()
        self._set_running(True)
        self._clear_log()
        self.log(f"URA Suite {VERSION}")      # 10.18: which build made this log (a 10.16 log was read as 10.17)
        self.log(f"Run started: {len(run_list)} client(s) x {len(mods)} module(s) "
                 f"= {len(plan)} job(s). Log file: {log_path}")
        for c, m in skipped:
            self.log(f"Skipped (already completed today): {c['name']} - {m.title}", "warn")

        def work():
            rows = []
            from core.session import open_session
            by_client = []

            def closed_note(ctx, current, later_modules, later_clients):
                """One clear line instead of a failure per remaining step (live log 6 Oct 2026, 20:36)."""
                skipped = [m.title for m in later_modules]
                if later_clients:
                    skipped.append(f"{later_clients} more client(s)")
                ctx.log(f"The browser window was closed during {current}"
                        + (f" - {', '.join(skipped)} not started." if skipped else ".")
                        + " Run again: whatever is already saved is skipped.", "warn")
            for c in run_list:
                mods_c = [m for cc, m in plan if cc is c]
                if mods_c:
                    by_client.append((c, mods_c))
            total = len(plan)
            done = 0
            try:
                for ci, (c, mods_c) in enumerate(by_client):
                    ctx.check()
                    ctx.client_label = f"{c['name']} ({c['login']})"
                    ctx.add_secret(c.get("password"))
                    self._asking_login = c["login"]
                    self.ui_q.put(("client_state", (c["login"], "working")))
                    rows_before = len(rows)
                    ctx.log(f"===== CLIENT {c['name']} ({c['login']}): "
                            + ", ".join(m.title for m in mods_c) + " =====")
                    ctx.after_client = []
                    try:
                        from core import net
                        net.wait_online(ctx.log, ctx.check)
                        with open_session(ctx, c) as page:
                            from core.session import go_home
                            for mi, m in enumerate(mods_c):
                                ctx.check()
                                done += 1
                                ctx.progress(done, total, f"{c['name']}\n{m.title}")
                                ctx.log(f"----- {m.title} -----")
                                ctx.set_scope(c["name"], c["login"], m.title)
                                started = datetime.datetime.now()
                                row = {"when": started.strftime("%d/%m/%Y %H:%M"), "client": c["name"],
                                       "login": c["login"], "module": m.title, "status": "Completed",
                                       "files": 0, "seconds": 0}
                                n_before = len(ctx.items)

                                def files_recorded():
                                    # 10.17: the Overview counts the same files the Files sheet lists for this
                                    # step (live 9 Oct 2026: Certificates showed 0 files after the PAYE error
                                    # although 15 WHT files were saved; Returns 41 against 42 in Files)
                                    return sum(1 for it in ctx.items[n_before:] if it["file"])
                                try:
                                    if mi > 0 and browser_gone(page):
                                        raise BrowserClosed(f"The browser window was closed before {m.title}")
                                    if mi > 0:          # previous module left the portal elsewhere
                                        go_home(ctx, page, c, before=m.title)
                                    n_before = len(ctx.items)
                                    res = m.run(ctx, c, options[m.id], page) or {}
                                    row.update({k: v for k, v in res.items()
                                                if isinstance(v, (int, str, bool))})
                                    row["files"] = files_recorded() or int(res.get("files") or 0)
                                    if res.get("skipped"):
                                        row["status"] = "Skipped"
                                    bad = sum(1 for it in ctx.items[n_before:]
                                              if needs_review(it["status"]))
                                    if bad and row["status"] == "Completed":
                                        row["status"] = f"Completed - {bad} item(s) to check"
                                except Stopped as e:
                                    row["status"] = "Stopped"
                                    row["seconds"] = int((datetime.datetime.now() - started).total_seconds())
                                    rows.append(row); store.add_history(row, ctx.secrets)
                                    if isinstance(e, BrowserClosed):
                                        closed_note(ctx, m.title, mods_c[mi + 1:], len(by_client) - ci - 1)
                                    raise
                                except Exception as e:
                                    if browser_gone(page, e):        # Chrome closed by hand (6 Oct 2026)
                                        row["status"] = "Stopped: browser closed"
                                        row["seconds"] = int((datetime.datetime.now() - started).total_seconds())
                                        rows.append(row); store.add_history(row, ctx.secrets)
                                        closed_note(ctx, m.title, mods_c[mi + 1:], len(by_client) - ci - 1)
                                        raise BrowserClosed(f"The browser window was closed during {m.title}") from e
                                    row["status"] = f"Error: {type(e).__name__}"
                                    row["note"] = ctx.scrub(e)[:300]
                                    row["files"] = files_recorded()       # what was saved before the error
                                    ctx.record(f"{m.title} (whole module)", "failed", str(e))
                                    ctx.log(f"{m.title} failed for {c['name']}: {e}", "err")
                                    ctx.log(traceback.format_exc().strip().splitlines()[-1], "err")
                                    # 10.15 (user: 'put Retry if anything else is failed'): the whole step
                                    # once more after the client's other steps; files already saved are skipped
                                    ctx.log(f"{m.title} is tried once more after this client's other steps.")
                                    ctx.after_client.append((f"{m.title} - second try",
                                                             lambda pg, m=m, c=c, o=options[m.id]:
                                                             whole_step_again(ctx, m, c, o, pg)))
                                row["seconds"] = int((datetime.datetime.now() - started).total_seconds())
                                rows.append(row); store.add_history(row, ctx.secrets)
                            if browser_gone(page):        # closed during the last step, which carried on
                                closed_note(ctx, mods_c[-1].title, [], len(by_client) - ci - 1)
                                raise BrowserClosed("The browser window was closed")
                            # 10.14: work a step left for the end of the client (returns slow to open);
                            # 10.15: also the second try of anything that failed (whole steps and items)
                            later, ctx.after_client = list(ctx.after_client), []
                            ctx.final_try = True        # updates the first try's items; nothing put off again
                            try:
                                for title, fn in later:
                                    ctx.check()
                                    ctx.log(f"----- {title} -----")
                                    ctx.set_scope(c["name"], c["login"], title)
                                    try:
                                        go_home(ctx, page, c, before=title)
                                        fn(page)
                                    except Stopped:
                                        raise
                                    except Exception as e:
                                        if browser_gone(page, e):
                                            closed_note(ctx, title, [], len(by_client) - ci - 1)
                                            raise BrowserClosed(f"The browser window was closed during {title}") from e
                                        ctx.record(f"{title}", "failed", str(e))
                                        ctx.log(f"{title} failed for {c['name']}: {e}", "err")
                            finally:
                                ctx.final_try = False
                                ctx.after_client = []
                    except Stopped:
                        self.ui_q.put(("client_state", (c["login"], "stopped")))
                        raise
                    except Exception as e:          # login / browser failure for this client
                        ctx.set_scope(c["name"], c["login"], "login")
                        ctx.record("login / browser", "failed", str(e))
                        ctx.log(f"Could not run {c['name']}: {e}", "err")
                        for m in mods_c:
                            rows.append({"when": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                                         "client": c["name"], "login": c["login"], "module": m.title,
                                         "status": f"Error: {type(e).__name__}", "note": ctx.scrub(e)[:300],
                                         "files": 0, "seconds": 0})
                            store.add_history(rows[-1], ctx.secrets)
                    mine = rows[rows_before:]
                    bad = any(not str(r.get("status", "")).startswith("Completed") for r in mine)
                    self.ui_q.put(("client_state", (c["login"], "failed" if bad else "done")))
                    # the client's Chrome has closed (open_session ended): say so (user, 6 Oct 2026)
                    voice.say(voice.phrase(ok=not bad, more=ci < len(by_client) - 1 and not self.stop_event.is_set()),
                              su.get("voice", voice.DEFAULT))
                ctx.log("--- All done. Ready for another run. ---", "ok")
                self.ui_q.put(("done", "Finished."))
            except BrowserClosed:
                ctx.log("--- Stopped: the browser window was closed. ---", "warn")
                voice.say(voice.phrase(stopped=True), su.get("voice", voice.DEFAULT))
            except Stopped:
                ctx.log("--- Stopped by you. ---", "warn")
                voice.say(voice.phrase(stopped=True), su.get("voice", voice.DEFAULT))
                self.ui_q.put(("done", "Stopped."))
            except BaseException as e:
                ctx.log(f"Could not run: {e}", "err")
                self.ui_q.put(("done", "Stopped with an error - see log."))
            finally:
                if su["summary_xlsx"] and rows:
                    try:
                        path = write_run_summary(rows, ctx.items, out_root / f"Run_Summary_{stamp}.xlsx",
                                                 master=out_root / "URA_Suite_Summary.xlsx",
                                                 run_stamp=stamp)
                        n_fail = sum(1 for i in ctx.items if needs_review(i["status"]))
                        n_saved = sum(1 for i in ctx.items if i["status"] == "saved")
                        ctx.log(f"Run summary: {path}  ({n_saved} file(s) saved; "
                                f"to check: {n_fail})  master: {out_root / 'URA_Suite_Summary.xlsx'}", "ok")
                    except Exception as e:
                        ctx.log(f"Could not write the run summary: {e}", "warn")
                if su["open_folder_when_done"] and rows:
                    self.ui_q.put(("open_folder", str(out_root)))
                self.ui_q.put(("review", (rows, list(ctx.items),
                                          str(out_root / f"Run_Summary_{stamp}.xlsx"))))
                ctx.close()

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def _preflight(self, out_root: Path, su: dict):
        """Checks that would otherwise fail at the END of a long run."""
        try:
            out_root.mkdir(parents=True, exist_ok=True)
            probe = out_root / ".write_test"
            probe.write_text("ok"); probe.unlink()
        except Exception as e:
            return (f"The output folder cannot be written to:\n{out_root}\n\n{e}\n\n"
                    "Choose another folder with 'Output folder' in the menu.")
        try:
            free_mb = shutil.disk_usage(out_root).free // (1024 * 1024)
        except Exception:
            free_mb = None
        if free_mb is not None and free_mb < su["min_free_mb"]:
            return (f"Only {free_mb} MB free on the drive of {out_root}; the limit in Settings "
                    f"is {su['min_free_mb']} MB.")
        return None

    def _apply_engine_settings(self, su: dict):
        from engines import ledger_engine as le
        le.BROWSER_CHANNEL = "" if su["browser"] == "auto" else su["browser"]
        le.PROXY = su["proxy"]
        factor = float(su.get("speed_factor", 1) or 1)
        le.SLOW_FACTOR = factor
        le.NAV_TIMEOUT_MS = int(int(su["nav_timeout_s"]) * 1000 * factor)
        from engines import wht_engine as we, paye_engine as pe, returns_common as rc
        we.RESULT_WAIT_MS = int(max(60000, int(su["nav_timeout_s"]) * 1000) * factor)
        we.RELOAD_GRACE_S = 6 * factor
        pe.WAIT_S = int(60 * factor)
        rc.TAB_WAIT_SECONDS = rc.PAGE_WAIT_SECONDS = int(60 * factor)
        # speed-ups are applied when the shared browser opens, so read the
        # Ledgers page's tick here, before any session starts
        try:
            le.SPEED_UPS = bool(self.ledger_speed_choice())
        except Exception:
            le.SPEED_UPS = True

    def ledger_speed_choice(self):
        for m in MODULES:
            if m.id == "ledger" and hasattr(m, "speed"):
                return m.speed.get()
        return True

    def stop(self):
        self.stop_event.set()
        self.status_var.set("Stopping after the current step...")
        self.log("Stop requested - finishing the current step.", "warn")

    def _set_running(self, running):
        st = "disabled" if running else "normal"
        T = self.T
        self.run_sel_btn.configure(state=st)
        self.run_one_btn.configure(state=st)
        self.stop_btn.configure(state="normal" if running else "disabled")
        _ = T
        self.status_var.set("Running..." if running else "Ready.")

    # ------------------------------------------------------------ queue -> UI (main thread only)
    def _pump(self):
        try:
            while True:
                kind, payload = self.ui_q.get_nowait()
                if kind == "log":
                    self.log(*payload)
                elif kind == "status":
                    self.status_var.set(payload)
                elif kind == "progress":
                    i, n, title, step = payload
                    self.menu_status.configure(text=f"Client {i} of {n}\n{title}\n{step}")
                elif kind == "ask":
                    self._asking_for = payload[3] if len(payload) > 3 else ""
                    self._show_question(*payload[:3])
                elif kind == "hide_ask":
                    self._hide_question()
                elif kind == "review":
                    self.review.render(*payload)
                    self.show(self.review)
                elif kind == "client_state":
                    self._set_client_state(*payload)
                elif kind == "open_folder":
                    try:
                        self.open_output()
                    except Exception:
                        pass
                elif kind == "done":
                    self._hide_question()
                    self._set_running(False)
                    self.status_var.set(payload)
                    self.menu_status.configure(text="")
        except queue.Empty:
            pass
        self.after(100, self._pump)

    def log(self, line, kind=""):
        if not line.startswith("["):
            line = f"[{datetime.datetime.now():%H:%M:%S}] {line}"
        self.logbox.configure(state="normal")
        self.logbox.insert("end", line + "\n", kind or ())
        self.logbox.see("end")
        self.logbox.configure(state="disabled")

    def _clear_log(self):
        self.logbox.configure(state="normal")
        self.logbox.delete("1.0", "end")
        self.logbox.configure(state="disabled")

    # ------------------------------------------------------------ question bar
    def _answer(self, value):
        self.reply_q.put(value)
        self._hide_question()

    def _show_question(self, message, kind, defaults):
        for w in self.ask_row.winfo_children():
            w.destroy()
        T = self.T
        who = getattr(self, "_asking_for", "")
        self.ask_who.configure(text=(f"Your turn - {who}" if who else "Your turn"))
        self.ask_msg.configure(text=message)
        pb = T["prompt_btn"]
        pfg = T["prompt_bg"] if pb.upper() == "#FFFFFF" else "#FFFFFF"

        def B(text, val, green=False):
            return WebButton(self.ask_row, text=text, command=lambda: self._answer(val),
                             kind="prompt" if green else "prompt_secondary",
                             surface=T["prompt_bg"])
        _ = (pb, pfg)
        if kind == "login":
            B("Log In  (security code typed)", "", True).pack(side="left")
            B("Cancel run", "q").pack(side="left", padx=6)
            self.bind("<Return>", lambda e: self._answer(""))
        elif kind == "dates":
            fv, tv = tk.StringVar(value=defaults[0]), tk.StringVar(value=defaults[1])
            tk.Label(self.ask_row, text="From (dd/mm/yyyy)", bg=T["prompt_bg"], fg=T["prompt_fg"],
                     font=T["FONT"]).pack(side="left")
            ttk.Entry(self.ask_row, textvariable=fv, width=12).pack(side="left", padx=4)
            tk.Label(self.ask_row, text="To", bg=T["prompt_bg"], fg=T["prompt_fg"],
                     font=T["FONT"]).pack(side="left")
            ttk.Entry(self.ask_row, textvariable=tv, width=12).pack(side="left", padx=4)

            def use():
                try:
                    a = datetime.datetime.strptime(fv.get().strip(), "%d/%m/%Y")
                    b = datetime.datetime.strptime(tv.get().strip(), "%d/%m/%Y")
                except ValueError:
                    messagebox.showerror("Dates", "Type both dates as dd/mm/yyyy."); return
                if a > b:
                    messagebox.showerror("Dates", "'From' is after 'To'."); return
                self._answer(("dates", fv.get().strip(), tv.get().strip()))
            WebButton(self.ask_row, text="Use these dates", command=use, kind="prompt",
                      surface=T["prompt_bg"]).pack(side="left", padx=4)
            B("I set them in the browser", ("", None, None)).pack(side="left", padx=4)
            B("Skip", ("s", None, None)).pack(side="left", padx=4)
            B("Stop this section", ("q", None, None)).pack(side="left", padx=4)
        elif kind == "skipwait":                  # the program keeps waiting while this shows (10.14)
            B("Skip this return", "s", True).pack(side="left")
            B("Keep waiting", "w").pack(side="left", padx=6)
        elif kind == "manual":
            B("Save this page for Claude", "capture", True).pack(side="left")
            B("Close browser", "q").pack(side="left", padx=6)
        else:
            B("Done - continue", "", True).pack(side="left")
            B("Skip", "s").pack(side="left", padx=6)
            B("Stop this section", "q").pack(side="left")
            self.bind("<Return>", lambda e: self._answer(""))
        self.ask.pack(fill="x", padx=12, pady=(4, 6), before=self.log_frame)
        self.status_var.set("Waiting for you.")
        login = getattr(self, "_asking_login", "")
        if login:
            self._set_client_state(login, "your turn")
        self.deiconify(); self.lift()
        self.attributes("-topmost", True)
        self.after(1200, lambda: self.attributes("-topmost", False))
        self.bell()

    def _hide_question(self):
        self.ask.pack_forget()
        self.unbind("<Return>")
        login = getattr(self, "_asking_login", "")
        if login and self.client_states.get(login) == "your turn":
            self._set_client_state(login, "working")

    def start_test_login(self, client):
        from modules.test_login import test_login
        su = suite_settings(self.settings)
        self._apply_engine_settings(su)
        self.stop_event.clear()
        log_path = store.LOG_DIR / f"testlogin_{datetime.datetime.now():%Y%m%d_%H%M%S}.txt"
        ctx = RunContext(self.ui_q, self.reply_q, self.stop_event,
                         Path(self.settings["output_root"]), log_path)
        ctx.client_label = f"{client['name']} ({client['login']})"
        self._set_running(True)
        self.log(f"Test login for {client['name']} ({client['login']})...")

        def work():
            try:
                ok, detail = test_login(ctx, client)
                ctx.log(("LOGIN OK - " if ok else "LOGIN FAILED - ") + detail, "ok" if ok else "err")
                self.ui_q.put(("done", "Test finished."))
            except Stopped:
                ctx.log("Test login cancelled.", "warn"); self.ui_q.put(("done", "Cancelled."))
            except BaseException as e:
                ctx.log(f"Test login error: {e}", "err"); self.ui_q.put(("done", "Test failed."))
            finally:
                ctx.close()
        self.worker = threading.Thread(target=work, daemon=True); self.worker.start()

    def open_manual(self):
        """Log in as one client and leave the browser to the person."""
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Busy", "Wait for the current run to finish."); return
        sel = self.tree.selection()
        login = self.tree.item(sel[0], "values")[2] if sel else None
        client = next((c for c in self.clients if c["login"] == login), None) or \
            next((c for c in self.clients if c["login"] in self.ticked_logins), None)
        if client is None:
            messagebox.showinfo("Open portal", "Add a client first (Clients), then select it in the list.")
            return
        if not messagebox.askokcancel("Open portal (manual)",
                                      f"Log in as {client['name']} ({client['login']}) and leave the "
                                      "browser open for you to use?\n\nNothing is clicked or downloaded "
                                      "by the program. Press 'Close browser' in this window when you "
                                      "are finished."):
            return
        from core.manual import run_manual
        from core.session import open_session
        su = suite_settings(self.settings)
        self._apply_engine_settings(su)
        self.stop_event.clear()
        log_path = store.LOG_DIR / f"manual_{datetime.datetime.now():%Y%m%d_%H%M%S}.txt"
        ctx = RunContext(self.ui_q, self.reply_q, self.stop_event,
                         Path(self.settings["output_root"]), log_path)
        ctx.client_label = f"{client['name']} ({client['login']})"
        ctx.login_timeout_s = su["login_timeout_s"]
        self._asking_login = client["login"]
        self._set_running(True)
        self.log(f"Opening the portal for {client['name']} ({client['login']}) - manual mode.")

        def work():
            try:
                run_manual(ctx, client, open_session)
                self.ui_q.put(("done", "Browser closed."))
            except Stopped:
                ctx.log("Manual browser closed.", "warn"); self.ui_q.put(("done", "Closed."))
            except BaseException as e:
                ctx.log(f"Manual browser error: {e}", "err"); self.ui_q.put(("done", "Error - see log."))
            finally:
                ctx.close()
        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def _on_close(self):
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("Still running", "A download is running. Stop it and close?"):
                return
            self.stop_event.set()
        self.destroy()


def whole_step_again(ctx, module, client, options, page):
    """10.15: a step that failed as a whole, run once more at the end of the client (ctx.final_try is set, so
    its items update the first try's and nothing is put off again).  Files already saved are skipped by each
    step's own resume rules.  When it ends without an error, the first try's '(whole module)' item says so."""
    res = module.run(ctx, client, options, page) or {}
    ctx.resolve(f"{module.title} (whole module)", "done on second try",
                "failed in the first try; the second try at the end of the client finished")
    ctx.log(f"{module.title}: second try finished.", "ok")
    return res


def default_period():
    """Current Ugandan financial year (July-June) up to this month."""
    t = datetime.date.today()
    start = t.year if t.month >= 7 else t.year - 1
    return f"{start}-07", f"{t.year}-{t.month:02d}"


class ReviewPage:
    """Briefing step 10: what the last run did, and what needs a person."""
    id = "review"
    title = "Review"

    def __init__(self, app):
        self.app = app
        self.summary_path = None

    def build_panel(self, parent):
        T = self.app.T
        f = ttk.Frame(parent, padding=10)
        ttk.Label(f, text="Review", style="Head.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(f, text="What the last run did, and every item that needs a person to check it "
                          "before the client's file is finalised.", style="Muted.TLabel").grid(
            row=1, column=0, sticky="w", pady=(0, 8))
        self.head = ttk.Label(f, text="No run yet in this session.", font=T["FONT_BOLD"])
        self.head.grid(row=2, column=0, sticky="w")
        cols = ("client", "step", "status", "files", "time")
        self.res = ttk.Treeview(f, columns=cols, show="headings", height=4)
        for c, h, w in zip(cols, ("Client", "Step", "Status", "Files", "Minutes"), (200, 150, 260, 60, 70)):
            self.res.heading(c, text=h); self.res.column(c, width=w, anchor="w")
        self.res.grid(row=3, column=0, sticky="ew", pady=(4, 8))
        ttk.Label(f, text="Needs your review", font=T["FONT_SUB"]).grid(row=4, column=0, sticky="w")
        cols2 = ("client", "step", "item", "reason")
        self.todo = ttk.Treeview(f, columns=cols2, show="headings", height=4)
        for c, h, w in zip(cols2, ("Client", "Step", "Item", "Reason"), (160, 110, 300, 330)):
            self.todo.heading(c, text=h); self.todo.column(c, width=w, anchor="w")
        self.todo.grid(row=5, column=0, sticky="ew", pady=(4, 8))
        bar = ttk.Frame(f)
        bar.grid(row=6, column=0, sticky="w")
        WebButton(bar, text="Open the run summary (Excel)", command=self.open_summary,
                  kind="primary").pack(side="left", padx=(0, 6))
        WebButton(bar, text="Open the output folder", command=self.app.open_output).pack(side="left")
        f.columnconfigure(0, weight=1)
        return f

    def render(self, rows, items, path):
        self.summary_path = path
        for t in (self.res, self.todo):
            t.delete(*t.get_children())
        for r in rows:
            self.res.insert("", "end", values=(r.get("client", ""), r.get("module", ""), r.get("status", ""),
                                               r.get("files", ""), round(int(r.get("seconds") or 0) / 60, 1)))
        bad = [i for i in items if needs_review(i["status"])]
        for i in bad:
            self.todo.insert("", "end", values=(i["client"], i["module"], i["item"][:80], i["detail"][:120]))
        saved = sum(1 for i in items if i["status"] == "saved")
        self.head.configure(text=f"{len(rows)} step(s) run - {saved} file(s) saved - "
                                 f"{len(bad)} item(s) need your review"
                                 + ("" if bad else " - nothing flagged"))

    def open_summary(self):
        p = self.summary_path
        if p and os.path.exists(p):
            if sys.platform.startswith("win"):
                os.startfile(p)
            else:
                subprocess.Popen(["xdg-open", p])
        else:
            messagebox.showinfo("Review", "No run summary yet - run at least one step first.")


# ============================================================ run summary
def _style_header(ws, ncols):
    from openpyxl.styles import Font, PatternFill
    for c in ws[1][:ncols]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79")
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


def _fill_sheet(ws, header, rows, widths, link_col=None):
    from openpyxl.styles import Font, PatternFill
    from engines.sheet_safety import neutralise_sheet
    ws.append(header)
    for r in rows:
        ws.append(r)
    neutralise_sheet(ws)                    # client names and portal text stay text (L2)
    _style_header(ws, len(header))
    for col, w in zip("ABCDEFGHIJKL", widths):
        ws.column_dimensions[col].width = w
    red = PatternFill("solid", fgColor="FDE0E0")
    amber = PatternFill("solid", fgColor="FFF3CC")
    si = header.index("Status") if "Status" in header else None
    for row in ws.iter_rows(min_row=2):
        if si is not None:
            st = str(row[si].value or "").lower()
            if needs_review(st):
                for c in row:
                    c.fill = red
            elif "stopped" in st or "skipped" in st:
                for c in row:
                    c.fill = amber
        if link_col is not None and row[link_col].value:
            try:
                row[link_col].hyperlink = Path(str(row[link_col].value)).resolve().as_uri()
            except Exception:
                row[link_col].hyperlink = None
            row[link_col].font = Font(color="0563C1", underline="single")


def write_run_summary(rows, items, path: Path, master: Path = None, run_stamp=""):
    """Three sheets: Overview (client x module), Files (every file produced,
    with a clickable link), Failed (every item that failed / was incomplete).
    If `master` is given, the same rows are appended to that workbook so one
    file holds every run."""
    from openpyxl import Workbook, load_workbook
    from engines.sheet_safety import neutralise_sheet
    ov_head = ["When", "Client", "Login", "Module", "Status", "Files", "Seconds", "Note"]
    ov_rows = [[r.get(k, "") for k in ("when", "client", "login", "module", "status", "files",
                                       "seconds", "note")] for r in rows]
    fi_head = ["Client", "Login", "Module", "Item", "Status", "Detail", "File"]
    fi_rows = [[i["client"], i["login"], i["module"], i["item"], i["status"], i["detail"], i["file"]]
               for i in items if i["file"]]
    fa_head = ["Client", "Login", "Module", "Item", "Status", "Reason"]
    fa_rows = [[i["client"], i["login"], i["module"], i["item"], i["status"], i["detail"]]
               for i in items if needs_review(i["status"])]
    other_rows = [[i["client"], i["login"], i["module"], i["item"], i["status"], i["detail"]]
                  for i in items if not i["file"] and i["status"].lower() not in ("saved",)
                  and not needs_review(i["status"])]

    wb = Workbook()
    ws = wb.active; ws.title = "Overview"
    _fill_sheet(ws, ov_head, ov_rows, (17, 28, 14, 24, 22, 8, 9, 50))
    _fill_sheet(wb.create_sheet("Files"), fi_head, fi_rows, (26, 13, 22, 44, 10, 24, 70), link_col=6)
    _fill_sheet(wb.create_sheet("Failed"), fa_head, fa_rows, (26, 13, 22, 50, 12, 60))
    _fill_sheet(wb.create_sheet("Other outcomes"), fa_head, other_rows, (26, 13, 22, 50, 14, 60))
    wb.save(path)

    if master is not None:
        try:
            if master.exists():
                mb = load_workbook(master)
            else:
                mb = Workbook(); mb.remove(mb.active)
            for name, head, data in (("Overview", ["Run"] + ov_head, ov_rows),
                                     ("Files", ["Run"] + fi_head, fi_rows),
                                     ("Failed", ["Run"] + fa_head, fa_rows)):
                ws = mb[name] if name in mb.sheetnames else mb.create_sheet(name)
                if ws["A1"].value is None:
                    for j, h in enumerate(head, 1):
                        ws.cell(row=1, column=j, value=h)
                for r in data:
                    ws.append([run_stamp] + r)
                neutralise_sheet(ws)
                _style_header(ws, len(head))
            mb.save(master)
        except Exception:
            pass                                    # the per-run file is already written
    return path


# ================================================================== dialogs
class ClientsDialog(tk.Toplevel):
    def __init__(self, app: App):
        super().__init__(app)
        self.app = app
        T = app.T
        self.title("Clients")
        from core import display
        display.place(self, 980, 440, 640, 320)
        self.configure(bg=T["card_bg"])
        self.transient(app); self.grab_set()
        tk.Label(self, text="Clients", bg=T["card_bg"], fg=T["head"], font=T["FONT_HEAD"],
                 anchor="w").pack(fill="x", padx=16, pady=(14, 2))
        tk.Label(self, text="Passwords are encrypted with your Windows account.", bg=T["card_bg"],
                 fg=T["muted"], font=T["FONT_SMALL"], anchor="w").pack(fill="x", padx=16)
        self.lb = tk.Listbox(self, font=FONT, bd=0, highlightthickness=1,
                             highlightbackground=T["line"], highlightcolor=T["accent"],
                             selectbackground=T["accent"], selectforeground=T["accent_fg"],
                             activestyle="none")
        self.lb.pack(fill="both", expand=True, padx=16, pady=10)
        bb = tk.Frame(self, bg=T["card_bg"]); bb.pack(fill="x", padx=14, pady=(0, 14))
        for t, c, kind in (("Add client", self.add, "primary"), ("Edit", self.edit, "secondary"),
                           ("Remove", self.remove, "danger"),
                           ("Test login", self.test_login, "secondary"),
                           ("Check old files for passwords", self.check_leaks, "secondary"),
                           ("Import from Excel", self.import_excel, "secondary"),
                           ("Excel template", self.excel_template, "secondary"),
                           ("Tick all", lambda: self._tick(True), "secondary"),
                           ("Untick all", lambda: self._tick(False), "secondary")):
            WebButton(bb, text=t, command=c, kind=kind, surface=T["card_bg"]).pack(side="left", padx=2)
        WebButton(bb, text="Close", command=self.destroy, surface=T["card_bg"]).pack(side="right", padx=2)
        self._fill()

    def _fill(self):
        self.lb.delete(0, "end")
        for c in self.app.clients:
            need = "   - password must be typed again" if not c.get("password") else ""
            self.lb.insert("end", f"  {c['name']}   ({c['login']}){need}")

    def _sel(self):
        s = self.lb.curselection()
        return s[0] if s else None

    def _form(self, existing=None):
        T = self.app.T
        d = tk.Toplevel(self); d.title("Client" if existing else "Add client")
        d.configure(bg=T["card_bg"]); d.transient(self); d.grab_set()
        d.resizable(False, False)
        vals = {k: tk.StringVar(value=(existing or {}).get(k, "")) for k in ("name", "login", "password")}
        tk.Label(d, text="Edit client" if existing else "Add a client", bg=T["card_bg"], fg=T["head"],
                 font=T["FONT_HEAD"], anchor="w").grid(row=0, column=0, sticky="w", padx=20, pady=(16, 8))
        fields = {}
        for i, (k, lbl) in enumerate((("name", "Client name"), ("login", "Login ID (TIN)"),
                                      ("password", "Password"))):
            tk.Label(d, text=lbl, bg=T["card_bg"], fg=T["ink"], font=T["FONT_BOLD"],
                     anchor="w").grid(row=1 + 2 * i, column=0, sticky="w", padx=20, pady=(6, 2))
            cls = PasswordField if k == "password" else TextField
            fields[k] = cls(d, textvariable=vals[k], width=34, surface=T["card_bg"])
            fields[k].grid(row=2 + 2 * i, column=0, sticky="w", padx=20)
        fields["name"].focus_set()
        out = {}
        err = tk.Label(d, text="", bg=T["card_bg"], fg=T["danger"], font=T["FONT_SMALL"], anchor="w")
        err.grid(row=7, column=0, sticky="w", padx=20, pady=(6, 0))

        def ok():
            v = {k: s.get().strip() for k, s in vals.items()}
            others = [c for c in self.app.clients if c is not existing]
            problem = store.client_problem(v, others)
            if problem:
                err.configure(text=problem)
                return
            out.update(v); d.destroy()
        bar = tk.Frame(d, bg=T["card_bg"]); bar.grid(row=8, column=0, sticky="e", padx=18, pady=(10, 16))
        WebButton(bar, text="Cancel", command=d.destroy, surface=T["card_bg"]).pack(side="left", padx=2)
        WebButton(bar, text="Save client", command=ok, kind="primary",
                  surface=T["card_bg"]).pack(side="left", padx=2)
        d.bind("<Return>", lambda e: ok())
        self._last_form = (d, fields)          # for tests
        d.wait_window()
        return out or None

    def add(self):
        v = self._form()
        if v:
            self.app.clients.append(v); self.app.ticked_logins.add(v["login"]); self._save()

    def edit(self):
        i = self._sel()
        if i is None:
            return
        v = self._form(self.app.clients[i])
        if v:
            self.app.clients[i] = v; self._save()

    def remove(self):
        i = self._sel()
        if i is None:
            return
        if messagebox.askyesno("Remove", f"Remove {self.app.clients[i]['name']}?", parent=self):
            self.app.ticked_logins.discard(self.app.clients[i]["login"])
            del self.app.clients[i]; self._save()

    def test_login(self):
        """Opens the portal, fills TIN + password, waits for the security code,
        reports whether the login worked, then closes. Nothing is downloaded."""
        i = self._sel()
        if i is None:
            messagebox.showinfo("Test login", "Select a client first.", parent=self); return
        if self.app.worker and self.app.worker.is_alive():
            messagebox.showinfo("Test login", "Wait for the current run to finish.", parent=self); return
        client = self.app.clients[i]
        self.destroy()
        self.app.start_test_login(client)

    def check_leaks(self):
        """Versions before 10.1 could write a password into logs, history or summaries (audit H6)."""
        from core import leakcheck
        args = (self.app.clients, store.LOG_DIR, store.HISTORY_FILE, Path(self.app.settings["output_root"]))
        found = leakcheck.scan(*args)
        if not found:
            messagebox.showinfo("Passwords in old files", "No saved password was found in the log folder, "
                                "the history or the summary workbooks.", parent=self)
            return
        lines = "\n".join(f"  - {Path(f).name}: {', '.join(n)}" for f, n in found[:25])
        more = f"\n  ... and {len(found) - 25} more" if len(found) > 25 else ""
        if messagebox.askyesno("Passwords in old files",
                               f"Saved passwords were found in {len(found)} file(s):\n\n{lines}{more}\n\n"
                               "Replace them with *** now? (Afterwards, consider changing those clients' "
                               "passwords if the files were ever shared.)", parent=self):
            leakcheck.scan(*args, redact=True)
            left = leakcheck.scan(*args)
            messagebox.showinfo("Passwords in old files",
                                "Done - none left." if not left else
                                f"{len(left)} file(s) could not be changed (open in Excel?) - close them and try again.",
                                parent=self)

    def excel_template(self):
        from core import client_import
        path = filedialog.asksaveasfilename(parent=self, title="Save the client template",
                                            defaultextension=".xlsx", initialfile="Clients_template.xlsx",
                                            filetypes=[("Excel workbook", "*.xlsx")])
        if path:
            client_import.make_template(path)
            messagebox.showinfo("Excel template", f"Saved: {path}\n\nFill one client per row, then use "
                                "'Import from Excel'.", parent=self)

    def import_excel(self):
        from core import client_import
        path = filedialog.askopenfilename(parent=self, title="Choose the Excel file with clients",
                                          filetypes=[("Excel workbook", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            res = client_import.read_clients(path, self.app.clients)
        except Exception as e:
            messagebox.showerror("Import from Excel", f"Could not read the file:\n{e}", parent=self)
            return
        text = client_import.summary(res)
        if not res.new and not res.updates:
            messagebox.showinfo("Import from Excel", "Nothing to import.\n\n" + text, parent=self)
            return
        update = False
        if res.updates:
            ans = messagebox.askyesnocancel("Import from Excel", text + "\n\nYes = add the new clients AND "
                                            "update the ones already in the list\nNo = add the new clients only"
                                            "\nCancel = import nothing", parent=self)
            if ans is None:
                return
            update = ans
        elif not messagebox.askokcancel("Import from Excel", text + "\n\nAdd them now?", parent=self):
            return
        self.app.clients = client_import.apply(self.app.clients, res, update_existing=update)
        for c in res.new:
            self.app.ticked_logins.add(c["login"])
        self._save()
        messagebox.showinfo("Import from Excel", f"Done: {len(res.new)} added"
                            + (f", {len(res.updates)} updated" if update else "") + ".\n\nThe Excel file "
                            "holds passwords in plain text - delete it or keep it somewhere safe.", parent=self)

    def _tick(self, val):
        self.app.ticked_logins = {c["login"] for c in self.app.clients} if val else set()
        self.app._refresh_clients()

    def _save(self):
        store.save_clients(self.app.clients)
        self._fill(); self.app._refresh_clients()


class HistoryDialog(tk.Toplevel):
    def __init__(self, app: App):
        super().__init__(app)
        self.title("Run history"); self.transient(app)
        from core import display
        display.place(self, 820, 420, 560, 300)
        cols = ("when", "client", "login", "module", "status", "files")
        t = ttk.Treeview(self, columns=cols, show="headings")
        for c, w in zip(cols, (130, 200, 120, 160, 140, 60)):
            t.heading(c, text=c.capitalize()); t.column(c, width=w)
        t.pack(fill="both", expand=True, padx=10, pady=10)
        for h in reversed(store.load_history()):
            t.insert("", "end", values=tuple(h.get(c, "") for c in cols))


if __name__ == "__main__":
    App().mainloop()
