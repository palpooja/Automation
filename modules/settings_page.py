"""
modules/settings_page.py - the Settings page in the vertical menu.

Not a run module: it has no tick box and no run().  Values are saved in
%LOCALAPPDATA%\\URA_Suite\\settings.json under "suite" and applied to the
engines by app.py before every run.
"""
import tkinter as tk
from tkinter import ttk

from core import store

DEFAULT_SUITE = {
    "browser": "auto",            # auto | chrome | msedge
    "proxy": "",
    "nav_timeout_s": 45,          # wait for a page element (Ledger engine)
    "login_timeout_s": 420,       # how long the yellow bar waits for you
    "skip_done_today": True,      # resume: skip client+module done today
    "open_folder_when_done": True,
    "min_free_mb": 500,
    "summary_xlsx": True,
    "theme": "revenue",           # ledger | revenue | runboard
    "speed_factor": 1,            # Internet speed: 1 normal, 2 slow, 4 very slow
    "log_keep_days": 30,          # screenshots / page copies / logs older than this are deleted at start-up
    "voice": "1",                 # 'Job completed' announcement: "1" / "4" (Voice_Samples.ps1) or "off"
}


def suite_settings(settings: dict) -> dict:
    s = dict(DEFAULT_SUITE)
    s.update(settings.get("suite", {}))
    return s


class SettingsPage:
    id = "settings"
    title = "Settings"

    def __init__(self, settings: dict, on_change):
        self.settings = settings
        self.on_change = on_change

    def build_panel(self, parent):
        cur = suite_settings(self.settings)
        f = ttk.Frame(parent, padding=10)
        ttk.Label(f, text="Settings", style="Head.TLabel").grid(row=0, column=0,
                                                                columnspan=2, sticky="w")
        ttk.Label(f, text="Applied to every run. Saved automatically.",
                  style="Muted.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 8))

        bf = ttk.LabelFrame(f, text=" Browser and network ", padding=8)
        bf.grid(row=2, column=0, sticky="ew")
        self.browser = tk.StringVar(value=cur["browser"])
        ttk.Label(bf, text="Browser").grid(row=0, column=0, sticky="w")
        ttk.Combobox(bf, textvariable=self.browser, state="readonly", width=28,
                     values=["auto", "chrome", "msedge"]).grid(row=0, column=1, sticky="w", pady=2)
        ttk.Label(bf, text="auto = Google Chrome, then Microsoft Edge, then the bundled copy",
                  style="Muted.TLabel").grid(row=1, column=0, columnspan=2, sticky="w")
        self.proxy = tk.StringVar(value=cur["proxy"])
        ttk.Label(bf, text="Proxy (blank = none)").grid(row=2, column=0, sticky="w")
        ttk.Entry(bf, textvariable=self.proxy, width=30).grid(row=2, column=1, sticky="w", pady=2)
        ttk.Label(bf, text="e.g. http://proxy.company.com:8080", style="Muted.TLabel").grid(
            row=3, column=0, columnspan=2, sticky="w")
        SPEEDS = {1: "Normal", 2: "Slow - every wait x2", 4: "Very slow - every wait x4"}
        self._speeds = SPEEDS
        self.speed = tk.StringVar(value=SPEEDS.get(int(cur.get("speed_factor", 1)), "Normal"))
        ttk.Label(bf, text="Internet speed").grid(row=6, column=0, sticky="w")
        ttk.Combobox(bf, textvariable=self.speed, state="readonly", width=28,
                     values=list(SPEEDS.values())).grid(row=6, column=1, sticky="w", pady=2)
        ttk.Label(bf, text="Slow / Very slow give the portal more time at every step. A dropped "
                           "connection is always waited out (checked before each retry).",
                  style="Muted.TLabel", wraplength=520, justify="left").grid(
            row=7, column=0, columnspan=2, sticky="w")
        self.nav = tk.IntVar(value=int(cur["nav_timeout_s"]))
        self.login = tk.IntVar(value=int(cur["login_timeout_s"]))
        ttk.Label(bf, text="Wait for a page element (seconds)").grid(row=4, column=0, sticky="w")
        ttk.Spinbox(bf, from_=15, to=300, textvariable=self.nav, width=6).grid(row=4, column=1, sticky="w", pady=2)
        ttk.Label(bf, text="Wait for you to type the security code (seconds)").grid(row=5, column=0, sticky="w")
        ttk.Spinbox(bf, from_=60, to=1800, textvariable=self.login, width=6).grid(row=5, column=1, sticky="w", pady=2)

        rf = ttk.LabelFrame(f, text=" Runs ", padding=8)
        rf.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        self.skip = tk.BooleanVar(value=cur["skip_done_today"])
        self.open = tk.BooleanVar(value=cur["open_folder_when_done"])
        self.summary = tk.BooleanVar(value=cur["summary_xlsx"])
        ttk.Checkbutton(rf, text="Resume: skip a client + module already completed today",
                        variable=self.skip).pack(anchor="w")
        ttk.Checkbutton(rf, text="Open the output folder when the run finishes",
                        variable=self.open).pack(anchor="w")
        ttk.Checkbutton(rf, text="Write Run_Summary_<date>.xlsx for every run",
                        variable=self.summary).pack(anchor="w")
        row = ttk.Frame(rf); row.pack(anchor="w", pady=(6, 0))
        self.free = tk.IntVar(value=int(cur["min_free_mb"]))
        ttk.Label(row, text="Refuse to start if free disk space is below (MB)").pack(side="left")
        ttk.Spinbox(row, from_=50, to=50000, increment=50, textvariable=self.free, width=7).pack(side="left", padx=6)
        row2 = ttk.Frame(rf); row2.pack(anchor="w", pady=(6, 0))
        self.keep = tk.IntVar(value=int(cur.get("log_keep_days", 30)))
        ttk.Label(row2, text="Delete screenshots, page copies and logs older than (days, 0 = never)").pack(side="left")
        ttk.Spinbox(row2, from_=0, to=3650, increment=1, textvariable=self.keep, width=5).pack(side="left", padx=6)
        af = ttk.Frame(rf)
        af.pack(anchor="w", pady=(8, 0))
        from core.theme import THEMES
        self._theme_ids = list(THEMES)
        labels = [THEMES[k]["label"] for k in self._theme_ids]
        cur_t = cur.get("theme", "revenue")
        self.theme = tk.StringVar(value=THEMES.get(cur_t, THEMES["revenue"])["label"])
        ttk.Label(af, text="Design").pack(side="left")
        ttk.Combobox(af, textvariable=self.theme, values=labels, state="readonly",
                     width=22).pack(side="left", padx=6)
        ttk.Label(af, text="applies the next time you start the program",
                  style="Muted.TLabel").pack(side="left")
        vf = ttk.Frame(rf)
        vf.pack(anchor="w", pady=(8, 0))
        from core import voice
        self._voice_ids = list(voice.CHOICES)
        self.voice = tk.StringVar(value=voice.CHOICES.get(cur.get("voice", voice.DEFAULT), "Sample 1"))
        ttk.Label(vf, text="Say 'Job completed' when a client's browser closes").pack(side="left")
        ttk.Combobox(vf, textvariable=self.voice, values=[voice.CHOICES[k] for k in self._voice_ids],
                     state="readonly", width=10).pack(side="left", padx=6)
        ttk.Button(vf, text="Test", command=lambda: voice.say(voice.phrase(ok=True, more=True),
                                                             self._voice_id())).pack(side="left")
        ttk.Label(rf, text=f"Clients are stored with {store.credential_store_kind()}.\n"
                           f"Settings and logs: {store.DATA_DIR}",
                  style="Muted.TLabel", justify="left").pack(anchor="w", pady=(10, 0))

        for v in (self.browser, self.proxy, self.nav, self.login, self.skip,
                  self.open, self.summary, self.free, self.theme, self.speed, self.keep, self.voice):
            v.trace_add("write", lambda *_: self._save())
        return f

    def _voice_id(self):
        from core import voice
        return next((k for k, v in voice.CHOICES.items() if v == self.voice.get()), voice.DEFAULT)

    def values(self) -> dict:
        def num(var, default):
            try:
                return int(var.get())
            except Exception:
                return default
        return {"browser": self.browser.get(), "proxy": self.proxy.get().strip(),
                "nav_timeout_s": num(self.nav, 45), "login_timeout_s": num(self.login, 420),
                "skip_done_today": self.skip.get(), "open_folder_when_done": self.open.get(),
                "summary_xlsx": self.summary.get(), "min_free_mb": num(self.free, 500),
                "theme": self._theme_id(), "log_keep_days": num(self.keep, 30), "voice": self._voice_id(),
                "speed_factor": next((k for k, v in self._speeds.items() if v == self.speed.get()), 1)}

    def _theme_id(self):
        from core.theme import THEMES
        for k, t in THEMES.items():
            if t["label"] == self.theme.get():
                return k
        return "revenue"

    def _save(self):
        self.settings["suite"] = self.values()
        store.save_settings(self.settings)
        if self.on_change:
            self.on_change(self.settings["suite"])
