"""
core/widgets.py - web-style controls drawn on a tkinter Canvas.

WebButton      rounded button: hover shade, PRESS MOVES the face down 2 px and
               the shadow underneath collapses; keyboard focus ring; Space /
               Enter activate it; greyed when disabled.
TextField      rounded input box; border turns the accent colour while typing.
PasswordField  TextField with an eye icon at the right: click (or Space/Enter
               on the icon) to show the password, click again to hide it.

Colours come from core.theme.CURRENT, so every design uses them unchanged.
The classes accept the same configure() keys the window code already uses
(text, state, command), so they drop in where tk.Button was.
"""
import tkinter as tk
import tkinter.font as tkfont

from core import theme

PRESS_DEPTH = 3          # px the face moves down when pressed
SHADOW = 4               # px of shadow under a resting button
RADIUS = 10


def _shade(hex_colour, factor):
    """factor < 1 darkens, > 1 lightens (towards white)."""
    h = hex_colour.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    if factor < 1:
        r, g, b = (int(c * factor) for c in (r, g, b))
    else:
        r, g, b = (int(c + (255 - c) * (factor - 1)) for c in (r, g, b))
    return f"#{max(0, min(255, r)):02X}{max(0, min(255, g)):02X}{max(0, min(255, b)):02X}"


def _round_rect(cv, x1, y1, x2, y2, r, **kw):
    r = max(1, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return cv.create_polygon(pts, smooth=True, **kw)


def _surface_of(parent, fallback):
    for key in ("bg", "background"):
        try:
            v = parent.cget(key)
            if v and v.startswith("#"):
                return v
        except Exception:
            continue
    return fallback


class WebButton(tk.Canvas):
    KINDS = ("primary", "secondary", "danger", "prompt", "prompt_secondary")

    def __init__(self, parent, text="", command=None, kind="secondary", font=None,
                 padx=16, pady=7, surface=None, width=None):
        T = theme.CURRENT
        self.T = T
        self._text, self._command, self._kind = text, command, kind
        self._state, self._hover, self._pressed, self._focus = "normal", False, False, False
        self._font = font or (T["FONT_BOLD"] if kind in ("primary", "prompt") else T["FONT"])
        f = tkfont.Font(font=self._font)
        from core.display import scale_of           # spacing grows with Windows display scaling
        s = scale_of(parent)
        padx, pady, width = int(padx * s), int(pady * s), (int(width * s) if width else width)
        self._padx, self._pady = padx, pady
        self._min_w = width or (f.measure(text) + 2 * padx)
        self._face_h = f.metrics("linespace") + 2 * pady
        surface = surface or _surface_of(parent, T["card_bg"])
        super().__init__(parent, width=self._min_w + 4, height=self._face_h + SHADOW + 4,
                         bg=surface, highlightthickness=0, bd=0, takefocus=1, cursor="hand2")
        for ev, fn in (("<Enter>", self._on_enter), ("<Leave>", self._on_leave),
                       ("<ButtonPress-1>", self._on_press), ("<ButtonRelease-1>", self._on_release),
                       ("<KeyPress-space>", self._on_key), ("<KeyPress-Return>", self._on_key),
                       ("<FocusIn>", lambda e: self._set("_focus", True)),
                       ("<FocusOut>", lambda e: self._set("_focus", False)),
                       ("<Configure>", lambda e: self._draw())):
            self.bind(ev, fn)
        self._draw()

    # ---- colours for the current kind/state
    def _colours(self):
        T = self.T
        if self._state == "disabled":
            return T["line"], T["muted"], None, T["line"]
        k = self._kind
        if k == "primary":
            face, text, border = T["accent"], T["accent_fg"], None
        elif k == "danger":
            face, text, border = T["card_bg"], T["danger"], T["danger"]
        elif k == "prompt":
            pb = T["prompt_btn"]
            face = pb
            text = T["prompt_bg"] if pb.upper() == "#FFFFFF" else "#FFFFFF"
            border = None
        elif k == "prompt_secondary":
            face, text, border = T["prompt_bg"], T["prompt_fg"], T["prompt_fg"]
        else:
            face, text, border = T["card_bg"], T["ink"], T["line"]
        if self._hover and not self._pressed:
            face = _shade(face, 0.92) if face.upper() != "#FFFFFF" else "#F1F4F7"
        if self._pressed:
            strong = k in ("primary", "prompt")
            face = _shade(face, 0.85 if strong else 0.95) if face.upper() != "#FFFFFF" else "#EEF2F5"
        shadow = _shade(face if k in ("primary", "prompt") else T["line"], 0.75)
        return face, text, border, shadow

    def _draw(self):
        self.delete("all")
        w = max(self.winfo_width(), self._min_w + 4)
        if w <= 4:
            w = self._min_w + 4
        face, text, border, shadow = self._colours()
        x1, x2 = 2, w - 2
        dy = PRESS_DEPTH if self._pressed else 0
        top = 2 + dy
        bottom = top + self._face_h
        if self._state != "disabled" and not self._pressed:
            _round_rect(self, x1, 2 + SHADOW, x2, 2 + SHADOW + self._face_h, RADIUS,
                        fill=shadow, outline="")
        elif self._pressed:
            _round_rect(self, x1, 2 + SHADOW, x2, 2 + SHADOW + self._face_h, RADIUS,
                        fill=_shade(shadow, 1.15), outline="")
        _round_rect(self, x1, top, x2, bottom, RADIUS, fill=face,
                    outline=border or face, width=1)
        if self._focus and self._state != "disabled":
            _round_rect(self, x1 - 1, top - 1, x2 + 1, bottom + 1, RADIUS + 1,
                        fill="", outline=self.T["accent"] if self._kind != "primary"
                        else _shade(self.T["accent"], 0.7), width=2)
        self.create_text(w / 2, (top + bottom) / 2, text=self._text, fill=text, font=self._font)

    # ---- events
    def _set(self, attr, val):
        setattr(self, attr, val)
        self._draw()

    def _on_enter(self, _e):
        if self._state != "disabled":
            self._set("_hover", True)

    def _on_leave(self, _e):
        self._hover = False
        self._set("_pressed", False)

    def _on_press(self, _e):
        if self._state == "disabled":
            return
        self.focus_set()
        self._set("_pressed", True)

    def _on_release(self, e):
        if self._state == "disabled" or not self._pressed:
            return
        inside = 0 <= e.x <= self.winfo_width() and 0 <= e.y <= self.winfo_height()
        self._set("_pressed", False)
        if inside and self._command:
            self._command()

    def _on_key(self, _e):
        if self._state == "disabled":
            return "break"
        self._set("_pressed", True)                     # same visible press as a click
        self.after(110, lambda: (self._set("_pressed", False),
                                 self._command() if self._command else None))
        return "break"

    def invoke(self):
        if self._state != "disabled" and self._command:
            self._command()

    # ---- tk.Button-compatible configuration
    def configure(self, cnf=None, **kw):
        if cnf:
            kw.update(cnf)
        redraw = False
        if "text" in kw:
            self._text = kw.pop("text")
            f = tkfont.Font(font=self._font)
            self._min_w = max(self._min_w, f.measure(self._text) + 2 * self._padx)
            tk.Canvas.configure(self, width=self._min_w + 4)
            redraw = True
        if "state" in kw:
            self._state = "disabled" if kw.pop("state") == "disabled" else "normal"
            tk.Canvas.configure(self, cursor="arrow" if self._state == "disabled" else "hand2")
            self._pressed = self._hover = False
            redraw = True
        if "command" in kw:
            self._command = kw.pop("command")
        if "kind" in kw:
            self._kind = kw.pop("kind")
            redraw = True
        for ignored in ("bg", "fg", "activebackground", "activeforeground", "font",
                        "disabledforeground", "relief", "bd", "padx", "pady",
                        "highlightthickness", "highlightbackground", "cursor"):
            kw.pop(ignored, None)
        if kw:
            tk.Canvas.configure(self, **kw)
        if redraw:
            self._draw()

    config = configure

    def cget(self, key):
        if key == "text":
            return self._text
        if key == "state":
            return self._state
        return tk.Canvas.cget(self, key)


class TextField(tk.Canvas):
    """Rounded input box.  .entry is the real tk.Entry inside."""

    def __init__(self, parent, textvariable=None, width=30, show="", surface=None,
                 right_pad=10):
        self._shown = getattr(self, "_shown", False)
        self._hot = getattr(self, "_hot", False)
        T = theme.CURRENT
        self.T = T
        self._font = T["FONT"]
        f = tkfont.Font(font=self._font)
        self._fw = f.measure("0") * width + 24 + right_pad
        self._fh = f.metrics("linespace") + 16
        surface = surface or _surface_of(parent, T["card_bg"])
        super().__init__(parent, width=self._fw, height=self._fh + 2, bg=surface,
                         highlightthickness=0, bd=0)
        self._focus = False
        self.entry = tk.Entry(self, textvariable=textvariable, show=show, bd=0, relief="flat",
                              highlightthickness=0, font=self._font, bg="#FFFFFF",
                              fg=T["ink"], insertbackground=T["ink"])
        self._right_pad = right_pad
        self._win = self.create_window(12, (self._fh + 2) / 2, window=self.entry, anchor="w",
                                       width=self._fw - 24 - right_pad)
        self.entry.bind("<FocusIn>", lambda e: self._set_focus(True))
        self.entry.bind("<FocusOut>", lambda e: self._set_focus(False))
        self.bind("<Button-1>", lambda e: self.entry.focus_set())
        self._draw()

    def _set_focus(self, v):
        self._focus = v
        self._draw()

    def _draw(self):
        self.delete("frame")
        T = self.T
        _round_rect(self, 1, 1, self._fw - 1, self._fh + 1, RADIUS, fill="#FFFFFF",
                    outline=T["accent"] if self._focus else T["line"],
                    width=2 if self._focus else 1, tags="frame")
        self.tag_lower("frame")
        self._draw_extra()

    def _draw_extra(self):
        pass

    def get(self):
        return self.entry.get()

    def focus_set(self):
        self.entry.focus_set()


class PasswordField(TextField):
    """Password box with an eye icon: shows / hides what you typed."""

    BULLET = "\u2022"

    def __init__(self, parent, textvariable=None, width=30, surface=None):
        self._shown = False
        self._hot = False
        self._tipwin = None
        super().__init__(parent, textvariable=textvariable, width=width, show=self.BULLET,
                         surface=surface, right_pad=34)
        self.configure(takefocus=1)
        # position-based, not tag-based: works on the very first click, no mouse move needed
        self.bind("<Button-1>", self._click)
        self.bind("<Motion>", lambda e: self._hover_eye(self._in_eye(e.x)))
        self.bind("<Leave>", lambda e: self._hover_eye(False))
        self.bind("<KeyPress-space>", lambda e: (self.toggle(), "break")[1])
        self.bind("<KeyPress-Return>", lambda e: (self.toggle(), "break")[1])
        self.bind("<FocusIn>", lambda e: self._draw())
        self.bind("<FocusOut>", lambda e: self._draw())

    @property
    def shown(self):
        return self._shown

    def _in_eye(self, x):
        return x >= self._fw - 40

    def _click(self, e):
        if self._in_eye(e.x):
            self.toggle()
        else:
            self.entry.focus_set()

    def _hover_eye(self, on):
        if on == self._hot:
            return
        self._hot = on
        tk.Canvas.configure(self, cursor="hand2" if on else "")
        self._draw()
        self._tip(on)

    def toggle(self):
        self._shown = not self._shown
        self.entry.configure(show="" if self._shown else self.BULLET)
        self._draw()
        if self._tipwin:
            self._tip(True)

    def _draw_extra(self):
        """Open eye while the password is hidden (click = show it);
        slashed eye while it is visible (click = hide it) - the Windows
        and common web convention."""
        self.delete("eye")
        T = self.T
        cx, cy = self._fw - 22, (self._fh + 2) / 2
        col = T["accent"] if self._hot else T["muted"]
        self.create_rectangle(cx - 15, cy - 13, cx + 15, cy + 13, fill="#FFFFFF", outline="",
                              tags="eye")                       # whole corner is clickable
        # almond outline as ONE smooth closed curve (thin arcs break up in Tk)
        self.create_line(cx - 11, cy, cx - 5, cy - 7, cx + 5, cy - 7, cx + 11, cy,
                         cx + 5, cy + 7, cx - 5, cy + 7, cx - 11, cy,
                         smooth=True, splinesteps=24, width=2, fill=col,
                         capstyle="round", joinstyle="round", tags="eye")
        self.create_oval(cx - 3.5, cy - 3.5, cx + 3.5, cy + 3.5, fill=col, outline="", tags="eye")
        if self._shown:
            self.create_line(cx - 10, cy + 9, cx + 10, cy - 9, fill="#FFFFFF", width=6,
                             capstyle="round", tags="eye")
            self.create_line(cx - 10, cy + 9, cx + 10, cy - 9, fill=col, width=2,
                             capstyle="round", tags="eye")
        try:
            focused = self.focus_get() is self
        except Exception:
            focused = False
        if focused:
            _round_rect(self, cx - 15, cy - 13, cx + 15, cy + 13, 6, fill="",
                        outline=T["accent"], width=2, tags="eye")

    def _tip(self, show):
        if getattr(self, "_tipwin", None):
            self._tipwin.destroy()
            self._tipwin = None
        if not show:
            return
        t = tk.Toplevel(self)
        t.wm_overrideredirect(True)
        x = self.winfo_rootx() + self._fw - 60
        y = self.winfo_rooty() + self._fh + 6
        t.wm_geometry(f"+{x}+{y}")
        tk.Label(t, text="Hide password" if self._shown else "Show password", bg="#2B2B2B",
                 fg="#FFFFFF", font=self.T["FONT_SMALL"], padx=8, pady=3).pack()
        self._tipwin = t
