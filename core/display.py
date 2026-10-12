"""core/display.py - make the window fit any screen size and Windows display scaling (6 Oct 2026).

Before 10.10 the window opened at a fixed 1240x820 and could not be made smaller than 1000x700, so on
a 1366x768 laptop, or a 1920x1080 screen at 150% scaling (which behaves like 1280x720), its lower part
was off the screen.  Here:
  * make_dpi_aware()  - Windows draws the program at the real screen sharpness, and Tk then sizes all
                        text in points correctly for the scaling (100 %, 125 %, 150 % ...);
  * scale_of(window)  - 1.0 at 96 DPI, 1.5 at 150 % - for sizes given in pixels;
  * place(window,...) - size and centre a window inside the visible screen;
  * ScrollArea        - a column that gets a scroll bar only when the screen is too short for it.
"""
import sys
import tkinter as tk
from tkinter import ttk

SCROLL_PASS = ("Text", "Treeview", "Listbox", "TCombobox", "Scrollbar", "TScrollbar")


def make_dpi_aware():
    """Call before the first Tk window.  Windows only; harmless elsewhere."""
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)        # Windows 8.1 and later
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()         # older Windows
        except Exception:
            pass


def scale_of(window) -> float:
    try:
        return max(1.0, window.winfo_fpixels("1i") / 96.0)
    except tk.TclError:
        return 1.0


def fit(screen_w, screen_h, scale, want_w, want_h, min_w, min_h):
    """(width, height, min width, min height, smaller than wanted) inside the visible screen.
    Leaves room for the title bar and the Windows taskbar."""
    avail_w = int(screen_w - 40 * scale)
    avail_h = int(screen_h - 90 * scale)
    w, h = min(int(want_w * scale), avail_w), min(int(want_h * scale), avail_h)
    return w, h, min(int(min_w * scale), w), min(int(min_h * scale), h), \
        (want_w * scale > avail_w or want_h * scale > avail_h)


def place(window, want_w, want_h, min_w, min_h, maximize_if_small=False):
    s = scale_of(window)
    sw, sh = window.winfo_screenwidth(), window.winfo_screenheight()
    w, h, mw, mh, small = fit(sw, sh, s, want_w, want_h, min_w, min_h)
    window.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 4)}")
    window.minsize(mw, mh)
    if small and maximize_if_small and sys.platform == "win32":
        try:
            window.state("zoomed")
        except tk.TclError:
            pass
    return s


class ScrollArea(tk.Frame):
    """Put widgets in .inner.  Widgets packed with expand=True still fill a tall screen; on a short
    screen the content keeps its natural height and a scroll bar appears."""

    def __init__(self, parent, bg):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.bar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hbar = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.bar.set, xscrollcommand=self.hbar.set)
        self.inner = tk.Frame(self.canvas, bg=bg)
        self._win = self.canvas.create_window(0, 0, window=self.inner, anchor="nw")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner.bind("<Configure>", self._sync, add="+")
        self.canvas.bind("<Configure>", self._sync, add="+")
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind_all(seq, self._wheel, add="+")

    def _sync(self, _event=None):
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return
        need = self.inner.winfo_reqheight()
        need_w = self.inner.winfo_reqwidth()
        height, width = max(ch, need), max(cw, need_w)
        self.canvas.itemconfigure(self._win, width=width, height=height)
        self.canvas.configure(scrollregion=(0, 0, width, height))
        wide = need_w > cw + 1                         # only on very narrow screens (about 1024 px)
        if wide and not self.hbar.winfo_ismapped():
            self.hbar.pack(side="bottom", fill="x", before=self.canvas)
        elif not wide and self.hbar.winfo_ismapped():
            self.hbar.pack_forget()
            self.canvas.xview_moveto(0)
        show = need > ch + 1
        if show and not self.bar.winfo_ismapped():
            self.bar.pack(side="right", fill="y", before=self.canvas)
        elif not show and self.bar.winfo_ismapped():
            self.bar.pack_forget()
            self.canvas.yview_moveto(0)

    def _wheel(self, event):
        if not self.bar.winfo_ismapped():
            return
        try:
            w = event.widget
            if not str(w).startswith(str(self)) or w.winfo_class() in SCROLL_PASS:
                return                                 # the log and lists scroll themselves
        except (tk.TclError, AttributeError):
            return
        if getattr(event, "num", None) in (4, 5):
            step = -1 if event.num == 4 else 1
        else:
            step = -1 if event.delta > 0 else 1
        self.canvas.yview_scroll(step * 3, "units")
