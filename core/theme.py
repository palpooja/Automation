"""
core/theme.py - the three visual designs and the bundled typefaces.

A design is only data (colours, fonts, a few layout switches); the window code
reads it once at start-up, so every design behaves identically.

Typefaces: the .ttf files in fonts/ are registered for THIS program only
(Windows AddFontResourceEx with FR_PRIVATE): nothing is installed into Windows
and they disappear when the program closes.  If registration fails the
window falls back to Segoe UI.
"""
import os
import sys
from pathlib import Path

THEMES = {
    "ledger": {
        "label": "A - Ledger sheet",
        "rail_bg": "#E6ECEA", "rail_fg": "#17242C", "rail_muted": "#55666C",
        "rail_active_bg": "#FBFCFB", "rail_active_fg": "#0E6E70", "rail_rule": "#C9D3D0",
        "rail_dark": False,
        "page_bg": "#FBFCFB", "card_bg": "#FBFCFB", "line": "#C9D3D0",
        "ink": "#17242C", "muted": "#5B6B70", "head": "#17242C", "frame_label": "#0E6E70",
        "accent": "#0E6E70", "accent_fg": "#FFFFFF",
        "prompt_bg": "#FFF1DC", "prompt_line": "#B86E00", "prompt_fg": "#5C3700",
        "prompt_btn": "#B86E00",
        "log_bg": "#F1F4F3", "log_fg": "#17242C", "log_ok": "#1F6B3A", "log_warn": "#8A5300",
        "log_err": "#A8322A", "danger": "#A8322A",
        "body": ["Atkinson Hyperlegible", "Segoe UI"], "head_font": ["Atkinson Hyperlegible", "Segoe UI"],
        "head_weight": "bold", "log_font": ["Cascadia Mono", "Consolas", "DejaVu Sans Mono"],
        "size": 10, "head_size": 15, "rail_size": 11,
        "queue_in_rail": False, "tab_join": True,
    },
    "revenue": {
        "label": "B - Revenue office",
        "rail_bg": "#12283D", "rail_fg": "#FFFFFF", "rail_muted": "#9DB0C4",
        "rail_active_bg": "#1C3A57", "rail_active_fg": "#FFFFFF", "rail_rule": "#2A4664",
        "rail_dark": True,
        "page_bg": "#EEF2F6", "card_bg": "#FFFFFF", "line": "#D5DEE8",
        "ink": "#14212E", "muted": "#5E738A", "head": "#12283D", "frame_label": "#12283D",
        "accent": "#12283D", "accent_fg": "#FFFFFF",
        "prompt_bg": "#0E6E70", "prompt_line": "#0E6E70", "prompt_fg": "#FFFFFF",
        "prompt_btn": "#FFFFFF",
        "log_bg": "#12283D", "log_fg": "#D9E3EE", "log_ok": "#8FD6A8", "log_warn": "#F2CF7A",
        "log_err": "#FF9C8F", "danger": "#B3261E",
        "body": ["Barlow", "Segoe UI"],
        "head_font": ["Barlow Condensed SemiBold", "Barlow Condensed", "Bahnschrift SemiCondensed", "Segoe UI Semibold"],
        "head_weight": "auto", "log_font": ["Cascadia Mono", "Consolas", "DejaVu Sans Mono"],
        "size": 10, "head_size": 18, "rail_size": 12,
        "queue_in_rail": False, "tab_join": False,
    },
    "runboard": {
        "label": "C - Run board",
        "rail_bg": "#0E3F44", "rail_fg": "#EAF3F3", "rail_muted": "#8FB3B6",
        "rail_active_bg": "#16565C", "rail_active_fg": "#FFFFFF", "rail_rule": "#245C61",
        "rail_dark": True,
        "page_bg": "#F4F7F7", "card_bg": "#FFFFFF", "line": "#D3DEDF",
        "ink": "#1B2B2E", "muted": "#5D7275", "head": "#0E3F44", "frame_label": "#0E3F44",
        "accent": "#0E6E70", "accent_fg": "#FFFFFF",
        "prompt_bg": "#FCEBC7", "prompt_line": "#D99A1B", "prompt_fg": "#4D3300",
        "prompt_btn": "#0E6E70",
        "log_bg": "#FFFFFF", "log_fg": "#1B2B2E", "log_ok": "#2F7D4E", "log_warn": "#9A6A06",
        "log_err": "#B23B2E", "danger": "#B23B2E",
        "body": ["IBM Plex Sans Condensed", "IBM Plex Sans Cond", "Segoe UI"],
        "head_font": ["IBM Plex Sans Condensed SemiBold", "IBM Plex Sans Cond SmBld",
                      "IBM Plex Sans Condensed", "Segoe UI Semibold"],
        "head_weight": "auto", "log_font": ["Cascadia Mono", "Consolas", "DejaVu Sans Mono"],
        "size": 10, "head_size": 16, "rail_size": 11,
        "queue_in_rail": True, "tab_join": False,
        "state_colours": {"waiting": "#8FB3B6", "working": "#EAF3F3", "your turn": "#F2B640",
                          "done": "#7FCB98", "failed": "#F08A7E", "stopped": "#F2B640",
                          "skipped": "#8FB3B6"},
    },
}
DEFAULT_THEME = "revenue"
CURRENT = None          # set by the window at start-up; read by core/widgets.py


def fonts_dir() -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / "fonts"


def register_fonts():
    """Make the bundled typefaces available to this process only."""
    d = fonts_dir()
    if not d.exists():
        return 0
    n = 0
    if os.name == "nt":
        try:
            import ctypes
            FR_PRIVATE = 0x10
            for f in d.glob("*.ttf"):
                n += ctypes.windll.gdi32.AddFontResourceExW(str(f), FR_PRIVATE, 0) > 0
        except Exception:
            pass
    return n


def pick(families, available):
    low = {a.lower(): a for a in available}
    for f in families:
        if f.lower() in low:
            return low[f.lower()]
    return families[-1]


def resolve(name, root):
    """Theme dict with concrete font tuples, for an existing Tk root."""
    import tkinter.font as tkfont
    t = dict(THEMES.get(name) or THEMES[DEFAULT_THEME])
    fams = tkfont.families(root)
    body = pick(t["body"], fams)
    head = pick(t["head_font"], fams)
    mono = pick(t["log_font"], fams)
    s = t["size"]
    w = t["head_weight"]
    if w == "auto":     # a family that IS the semibold cut (Windows names it so) needs no bold flag
        w = "normal" if any(k in head.lower() for k in ("semibold", "smbld")) else "bold"
    t["head_weight"] = w
    t["FONT"] = (body, s)
    t["FONT_SMALL"] = (body, s - 1)
    t["FONT_BOLD"] = (body, s, "bold")
    t["FONT_HEAD"] = (head, t["head_size"], t["head_weight"])
    t["FONT_SUB"] = (head, s + 2, t["head_weight"])
    t["FONT_RAIL"] = (body, t["rail_size"])
    t["FONT_BRAND"] = (head, t["head_size"] + 2, t["head_weight"])
    t["FONT_LOG"] = (mono, 9)
    t["families_used"] = (body, head, mono)
    return t
