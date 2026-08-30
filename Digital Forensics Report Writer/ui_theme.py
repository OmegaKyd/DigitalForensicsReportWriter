"""Shared forensic GUI theme used by the Digital Forensics Report Writer."""
import calendar as calmod
import os
import sys
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import ttk


def resource_dir():
    """Bundled files when frozen (PyInstaller), otherwise the script folder."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def writable_dir():
    """Folder next to the .exe (or the script folder) for settings and output."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

APP_NAME = "Digital Forensics Report Writer"
APP_VERSION = "1.0.1"
COPYRIGHT_YEAR = "2026"
GITHUB_URL = "https://github.com/omegakyd"
# Previous product name; used only to migrate AppData files from v1.0.0 and earlier.
LEGACY_APP_NAME = "Digital Evidence Report Writer"

# Palette aligned with Omega.png (navy field, cyan accent)
COLORS = {
    "bg": "#0b1e30",
    "panel": "#12283d",
    "card": "#16324a",
    "border": "#1e4a66",
    "accent": "#2ec8e0",
    "accent_dark": "#1788a0",
    "text": "#e7f3fb",
    "muted": "#8aa4b8",
    "entry_bg": "#0a1724",
    "entry_fg": "#e7f3fb",
    "button_bg": "#156f86",
    "button_fg": "#f4fcff",
    "header_bg": "#071524",
    "ok": "#3dce8a",
    "warn": "#e0b03a",
}


def _first_existing(paths):
    for candidate in paths:
        if candidate is not None and Path(candidate).exists():
            return str(candidate)
    return None


def logo_path():
    here = resource_dir()
    writable = writable_dir()
    return _first_existing((
        here / "assets" / "Omega.png",
        writable / "assets" / "Omega.png",
        here / "Omega.png",
        writable / "Omega.png",
        here.parent / "Omega.png",
    ))


def icon_paths():
    here = resource_dir()
    writable = writable_dir()
    ico = _first_existing((
        here / "assets" / "DFR_Writer.ico",
        writable / "assets" / "DFR_Writer.ico",
        here / "DFR_Writer.ico",
        writable / "DFR_Writer.ico",
    ))
    png = _first_existing((
        here / "assets" / "DFR_Writer.png",
        writable / "assets" / "DFR_Writer.png",
        here / "DFR_Writer.png",
        writable / "DFR_Writer.png",
    ))
    return (ico, png)


def apply_app_icon(root):
    """Set the window / taskbar icon (Omega + DFR Writer)."""
    ico_path, png_path = icon_paths()
    if ico_path:
        try:
            root.iconbitmap(ico_path)
        except Exception:
            pass
    if png_path:
        try:
            import tkinter as tk
            image = tk.PhotoImage(file=png_path)
            root.iconphoto(True, image)
            root._dfr_icon_image = image
        except Exception:
            pass


def parse_mdy_date(date_string, empty_ok=False):
    """Parse M/D/Y, MM/DD/YYYY, and dash variants. Works on Windows and Linux."""
    if not date_string or not isinstance(date_string, str) or not date_string.strip():
        if empty_ok:
            return ""
        raise ValueError("Date string cannot be empty")

    raw = date_string.strip()
    for sep in ("/", "-", "."):
        parts = raw.split(sep)
        if len(parts) != 3:
            continue
        month_s, day_s, year_s = (p.strip() for p in parts)
        if not (month_s.isdigit() and day_s.isdigit() and year_s.isdigit()):
            continue
        month, day, year = int(month_s), int(day_s), int(year_s)
        if year < 100:
            year += 2000
        try:
            date_obj = datetime(year, month, day)
        except ValueError:
            continue
        formatted = date_obj.strftime("%A, %B %d, %Y")
        formatted = formatted.replace(" 0", " ")
        return formatted

    raise ValueError(
        f"Invalid date format: {date_string}. Please use MM/DD/YYYY or M/D/YYYY."
    )


def format_mdy(value):
    if isinstance(value, datetime):
        value = value.date()
    return f"{value.month}/{value.day}/{value.year}"


def parse_entry_date(text):
    raw = (text or "").strip()
    for sep in ("/", "-", "."):
        parts = raw.split(sep)
        if len(parts) != 3 or not all(part.strip().isdigit() for part in parts):
            continue
        month, day, year = (int(part.strip()) for part in parts)
        if year < 100:
            year += 2000
        try:
            return date(year, month, day)
        except ValueError:
            break
    return None


def add_date_entry(parent, row, column, sticky="ew", pady=2):
    """Entry plus a compact calendar button in the same grid cell."""
    holder = ttk.Frame(parent)
    holder.grid(row=row, column=column, sticky=sticky, pady=pady)
    try:
        parent.columnconfigure(column, weight=1)
    except Exception:
        pass
    entry = ttk.Entry(holder)
    entry.pack(side="left", fill="x", expand=True)
    button = tk.Button(
        holder,
        text="📅",
        font=("Segoe UI Emoji", 9),
        relief="flat",
        bd=0,
        padx=5,
        pady=0,
        bg=COLORS["button_bg"],
        fg=COLORS["button_fg"],
        activebackground=COLORS["accent_dark"],
        activeforeground=COLORS["button_fg"],
        cursor="hand2",
        highlightthickness=0,
    )
    button.pack(side="left", padx=(3, 0))

    button.configure(command=lambda: entry.after(1, lambda: open_date_picker(entry, button)))
    return entry


def open_date_picker(entry, anchor=None):
    try:
        _open_date_picker(entry, anchor)
    except Exception as exc:
        try:
            from tkinter import messagebox
            messagebox.showerror("Calendar", f"Could not open the calendar:\n{exc}")
        except Exception:
            print(f"Could not open the calendar: {exc}")


def _widget_window(widget):
    """Use the visible report window, not the withdrawn start screen.

    Report windows set self.master to the start screen, so walking .master
    all the way up would place the calendar on the hidden start window.
    """
    current = widget
    last = widget
    while current is not None:
        last = current
        try:
            if current.winfo_class() in ("Tk", "Toplevel"):
                return current
        except Exception:
            pass
        current = getattr(current, "master", None)
    return last


def _close_date_picker(host):
    pop = getattr(host, "_date_picker", None)
    if pop is not None:
        try:
            pop.destroy()
        except Exception:
            pass
    host._date_picker = None


def _open_date_picker(entry, anchor=None):
    existing = parse_entry_date(entry.get())
    selected = existing or date.today()
    colors = COLORS
    host = _widget_window(entry)
    _close_date_picker(host)

    win = tk.Frame(host, bg=colors["border"], padx=1, pady=1, highlightthickness=0)
    inner = tk.Frame(win, bg=colors["panel"])
    inner.pack(fill="both", expand=True)
    host._date_picker = win

    month_var = tk.IntVar(master=host, value=selected.month)
    year_var = tk.IntVar(master=host, value=selected.year)

    header = tk.Frame(inner, bg=colors["panel"])
    header.pack(fill="x", padx=8, pady=(8, 4))
    tk.Button(header, text="<", width=3, relief="flat", command=lambda: shift(-1),
              bg=colors["button_bg"], fg=colors["button_fg"], bd=0).pack(side="left")
    title = tk.Label(header, text="", bg=colors["panel"], fg=colors["text"], font=("Segoe UI", 10, "bold"))
    title.pack(side="left", expand=True)
    tk.Button(header, text=">", width=3, relief="flat", command=lambda: shift(1),
              bg=colors["button_bg"], fg=colors["button_fg"], bd=0).pack(side="right")

    body = tk.Frame(inner, bg=colors["panel"])
    body.pack(padx=8, pady=(0, 8))

    def close():
        _close_date_picker(host)

    def choose(day_value, _event=None):
        picked = date(int(year_var.get()), int(month_var.get()), int(day_value))
        try:
            entry.focus_set()
            entry.delete(0, "end")
            entry.insert(0, format_mdy(picked))
            entry.event_generate("<FocusOut>")
        except Exception:
            pass
        close()
        return "break"

    def shift(delta_month):
        month = int(month_var.get()) + delta_month
        year = int(year_var.get())
        while month < 1:
            month += 12
            year -= 1
        while month > 12:
            month -= 12
            year += 1
        month_var.set(month)
        year_var.set(year)
        draw()

    def draw():
        for child in body.winfo_children():
            child.destroy()
        month = int(month_var.get())
        year = int(year_var.get())
        title.configure(text=f"{calmod.month_name[month]} {year}")
        for col, name in enumerate(("Su", "Mo", "Tu", "We", "Th", "Fr", "Sa")):
            tk.Label(body, text=name, width=3, bg=colors["panel"], fg=colors["muted"],
                     font=("Segoe UI", 8, "bold")).grid(row=0, column=col, padx=1, pady=1)
        weeks = calmod.Calendar(firstweekday=6).monthdayscalendar(year, month)
        today = date.today()
        for row_i, week in enumerate(weeks, start=1):
            for col, day_value in enumerate(week):
                if not day_value:
                    tk.Label(body, text="", width=3, bg=colors["panel"]).grid(row=row_i, column=col)
                    continue
                cell = date(year, month, day_value)
                is_selected = cell == selected
                is_today = cell == today
                bg = colors["accent_dark"] if is_selected else colors["entry_bg"]
                fg = colors["accent"] if is_today and not is_selected else colors["text"]
                day_lbl = tk.Label(
                    body,
                    text=str(day_value),
                    width=3,
                    relief="flat",
                    bd=0,
                    bg=bg,
                    fg=fg,
                    font=("Segoe UI", 9, "bold" if is_selected or is_today else "normal"),
                    cursor="hand2",
                )
                day_lbl.grid(row=row_i, column=col, padx=1, pady=1)
                day_lbl.bind("<ButtonRelease-1>", lambda e, d=day_value: choose(d, e))

    def choose_today():
        today = date.today()
        month_var.set(today.month)
        year_var.set(today.year)
        choose(today.day)

    def clear_date():
        entry.delete(0, "end")
        close()

    footer = tk.Frame(inner, bg=colors["panel"])
    footer.pack(fill="x", padx=8, pady=(0, 8))
    tk.Button(footer, text="Today", relief="flat", bd=0, padx=8,
              bg=colors["button_bg"], fg=colors["button_fg"], command=choose_today).pack(side="left")
    tk.Button(footer, text="Clear", relief="flat", bd=0, padx=8,
              bg=colors["button_bg"], fg=colors["button_fg"], command=clear_date).pack(side="left", padx=6)
    tk.Button(footer, text="Close", relief="flat", bd=0, padx=8,
              bg=colors["button_bg"], fg=colors["button_fg"], command=close).pack(side="right")

    draw()
    host.update_idletasks()
    widget = anchor if anchor is not None else entry
    try:
        x = widget.winfo_rootx() - host.winfo_rootx()
        y = widget.winfo_rooty() - host.winfo_rooty() + widget.winfo_height() + 2
    except Exception:
        x, y = 80, 80
    win.place(x=max(8, x), y=max(8, y))
    win.lift()
    win.focus_set()
    win.bind("<Escape>", lambda _e: close())


def size_window(root, width=1260, height=700, min_width=980, min_height=520):
    """Center a report window using the requested size, clamped to the screen."""
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    width = min(max(int(width), min_width), max(min_width, screen_width - 40))
    height = min(max(int(height), min_height), max(min_height, screen_height - 50))
    x = max(0, (screen_width - width) // 2)
    y = 24
    root.geometry(f"{width}x{height}+{x}+{y}")
    try:
        root.minsize(min_width, min_height)
    except Exception:
        pass


def close_and_return(app):
    """Close a report window and show the start screen again."""
    try:
        timer = getattr(app, "_save_timer", None)
        if timer:
            app.after_cancel(timer)
            app._save_timer = None
    except Exception:
        pass
    try:
        if hasattr(app, "_perform_auto_save"):
            app._perform_auto_save()
    except Exception:
        pass
    master = getattr(app, "master", None)
    try:
        app.destroy()
    except Exception:
        pass
    if not master:
        return
    try:
        apply_theme(master)
    except Exception:
        pass
    try:
        master.deiconify()
        master.lift()
        master.focus_force()
    except Exception:
        pass


def add_action_bar(parent):
    """Bottom bar that stays visible under the form."""
    import tkinter as tk
    from tkinter import ttk
    bar = ttk.Frame(parent)
    bar.pack(side=tk.BOTTOM, fill=tk.X, padx=12, pady=(4, 8))
    return bar


def pack_right_actions(*buttons):
    """Place action buttons as a group on the bottom-right, left-to-right order."""
    import tkinter as tk
    for button in reversed(buttons):
        if button is not None:
            button.pack(side=tk.RIGHT, padx=6)


def build_extracted_info_pane(parent, placeholder=""):
    """Uniform extracted-info box used on every exam window."""
    import tkinter as tk
    from tkinter import ttk
    from tkinter.scrolledtext import ScrolledText

    frame = ttk.LabelFrame(parent, text="Extracted File Information", padding="8")
    frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=(4, 2))
    text = ScrolledText(
        frame,
        wrap=tk.WORD,
        width=48,
        height=14,
        bg=COLORS["entry_bg"],
        fg=COLORS["text"],
        insertbackground=COLORS["accent"],
        relief="flat",
        borderwidth=0,
        highlightthickness=0,
    )
    text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
    if placeholder:
        text.insert("1.0", placeholder)
    text.config(state=tk.DISABLED)
    return frame, text


def apply_theme(root):
    """Apply the shared dark forensic ttk theme to a Tk window."""
    colors = COLORS
    try:
        root.configure(bg=colors["bg"])
    except Exception:
        pass
    apply_app_icon(root)

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure(".", background=colors["bg"], foreground=colors["text"], font=("Segoe UI", 10))
    style.configure("TFrame", background=colors["bg"])
    style.configure("Card.TFrame", background=colors["panel"])
    style.configure("HeaderBar.TFrame", background=colors["header_bg"])
    style.configure("TLabel", background=colors["bg"], foreground=colors["text"], font=("Segoe UI", 10))
    style.configure("Header.TLabel", background=colors["bg"], foreground=colors["accent"], font=("Segoe UI", 16, "bold"))
    style.configure("SubHeader.TLabel", background=colors["header_bg"], foreground=colors["text"], font=("Segoe UI", 13, "bold"))
    style.configure("Hint.TLabel", background=colors["panel"], foreground=colors["muted"], font=("Segoe UI", 9))
    style.configure("CategoryHeader.TLabel", background=colors["panel"], foreground=colors["accent"], font=("Segoe UI", 13, "bold"))
    style.configure("CategoryFrame.TFrame", background=colors["panel"], relief="solid", borderwidth=1)

    style.configure("TLabelframe", background=colors["panel"], foreground=colors["accent"], bordercolor=colors["border"], relief="solid", borderwidth=1)
    style.configure("TLabelframe.Label", background=colors["panel"], foreground=colors["accent"], font=("Segoe UI", 10, "bold"))

    style.configure("TButton", background=colors["button_bg"], foreground=colors["button_fg"], font=("Segoe UI", 10, "bold"), padding=(10, 6), bordercolor=colors["accent_dark"])
    style.map("TButton",
              background=[("active", colors["accent_dark"]), ("pressed", colors["accent"])],
              foreground=[("active", colors["button_fg"])])

    style.configure("Accent.TButton", background=colors["accent"], foreground=colors["header_bg"], font=("Segoe UI", 10, "bold"), padding=(12, 7))
    style.map("Accent.TButton", background=[("active", "#5ad7ea")])

    style.configure("TEntry", fieldbackground=colors["entry_bg"], foreground=colors["entry_fg"], insertcolor=colors["accent"], background=colors["entry_bg"])
    style.configure("TCombobox", fieldbackground=colors["entry_bg"], foreground=colors["entry_fg"], background=colors["entry_bg"], arrowcolor=colors["accent"])
    style.map("TCombobox",
              fieldbackground=[("readonly", colors["entry_bg"])],
              foreground=[("readonly", colors["entry_fg"])],
              background=[("readonly", colors["entry_bg"])])

    style.configure("TCheckbutton", background=colors["panel"], foreground=colors["text"], font=("Segoe UI", 10))
    style.configure("TRadiobutton", background=colors["panel"], foreground=colors["text"])
    style.configure("TScrollbar", background=colors["panel"], troughcolor=colors["bg"], arrowcolor=colors["accent"])
    style.configure("TPanedwindow", background=colors["bg"])
    style.configure("TSeparator", background=colors["border"])
    style.configure("Horizontal.TProgressbar", background=colors["accent"], troughcolor=colors["entry_bg"])

    return colors


def add_header_bar(parent, title, subtitle=None):
    """Add a compact header bar with the Omega logo when available."""
    import tkinter as tk

    colors = COLORS
    if not subtitle:
        subtitle = APP_NAME
    bar = tk.Frame(parent, bg=colors["header_bg"], height=64)
    bar.pack(fill=tk.X, side=tk.TOP)
    bar.pack_propagate(False)

    logo = logo_path()
    if logo:
        try:
            from PIL import Image, ImageTk
            img = Image.open(logo).resize((48, 48))
            photo = ImageTk.PhotoImage(img)
            lbl = tk.Label(bar, image=photo, bg=colors["header_bg"])
            lbl.image = photo
            lbl.pack(side=tk.LEFT, padx=(14, 10), pady=8)
        except Exception:
            try:
                photo = tk.PhotoImage(file=logo)
                # PhotoImage has no high-quality resize; use subsample if huge
                lbl = tk.Label(bar, image=photo, bg=colors["header_bg"])
                lbl.image = photo
                lbl.pack(side=tk.LEFT, padx=(14, 10), pady=4)
            except Exception:
                tk.Label(bar, text="Ω", bg=colors["header_bg"], fg=colors["accent"], font=("Segoe UI", 22, "bold")).pack(side=tk.LEFT, padx=(16, 8))
    else:
        tk.Label(bar, text="Ω", bg=colors["header_bg"], fg=colors["accent"], font=("Segoe UI", 22, "bold")).pack(side=tk.LEFT, padx=(16, 8))

    text_wrap = tk.Frame(bar, bg=colors["header_bg"])
    text_wrap.pack(side=tk.LEFT, fill=tk.Y, pady=8)
    tk.Label(text_wrap, text=title, bg=colors["header_bg"], fg=colors["text"], font=("Segoe UI", 13, "bold")).pack(anchor="w")
    tk.Label(text_wrap, text=subtitle, bg=colors["header_bg"], fg=colors["muted"], font=("Segoe UI", 9)).pack(anchor="w")

    tk.Label(bar, text=f"v{APP_VERSION}", bg=colors["header_bg"], fg=colors["muted"], font=("Segoe UI", 9)).pack(side=tk.RIGHT, padx=16)
    return bar


class FormTabs:
    """Persistent form tabs. Switching tabs hides pages; it does not clear fields."""

    def __init__(self, parent):
        import tkinter as tk
        from tkinter import ttk

        self.colors = COLORS
        self.header = tk.Frame(parent, bg=self.colors["bg"])
        self.header.pack(fill=tk.X, pady=(0, 6))
        self.body = ttk.Frame(parent)
        self.body.pack(fill=tk.BOTH, expand=True)
        self.pages = {}
        self.canvases = {}
        self.buttons = {}
        self.titles = {}
        self.complete = {}
        self.visited = set()
        self.selected = None

    def add_tab(self, key, title):
        import tkinter as tk
        from tkinter import ttk

        page_wrap = ttk.Frame(self.body)
        canvas = tk.Canvas(page_wrap, bg=self.colors["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(page_wrap, orient=tk.VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
        canvas.bind("<Configure>", lambda e, c=canvas, w=window_id: c.itemconfigure(w, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        btn = tk.Button(
            self.header,
            text=title,
            command=lambda k=key: self.show(k),
            relief="flat",
            bd=0,
            padx=14,
            pady=8,
            font=("Segoe UI", 10, "bold"),
            bg=self.colors["panel"],
            fg=self.colors["text"],
            activebackground=self.colors["accent_dark"],
            activeforeground=self.colors["text"],
            cursor="hand2",
        )
        btn.pack(side=tk.LEFT, padx=(0, 6))

        self.pages[key] = page_wrap
        self.canvases[key] = canvas
        self.buttons[key] = btn
        self.titles[key] = title
        self.complete[key] = False
        if self.selected is None:
            # Landing tab is already on screen; allow it to turn green
            # as soon as its required fields are filled.
            self.show(key, mark_visited=True)
        return inner

    def page(self, key):
        return self.pages[key]

    def show(self, key, mark_visited=True):
        if key not in self.pages:
            return
        if self.selected and self.selected in self.pages:
            self.pages[self.selected].pack_forget()
        self.selected = key
        if mark_visited:
            self.visited.add(key)
        self.pages[key].pack(fill="both", expand=True)
        self.refresh_styles()

    def set_complete(self, key, complete):
        if key not in self.buttons:
            return
        self.complete[key] = bool(complete)
        self.refresh_styles()

    def refresh_styles(self):
        for key, btn in self.buttons.items():
            title = self.titles.get(key, key)
            if self.complete.get(key) and (key in self.visited or key == self.selected):
                bg = self.colors["ok"]
                fg = self.colors["header_bg"]
                text = f"{title}  ✓"
            elif key == self.selected:
                bg = self.colors["accent_dark"]
                fg = self.colors["text"]
                text = title
            else:
                bg = self.colors["panel"]
                fg = self.colors["text"]
                text = title
            btn.configure(text=text, bg=bg, fg=fg, activebackground=bg, activeforeground=fg)

    def on_mousewheel(self, event):
        canvas = self.canvases.get(self.selected)
        if canvas is None:
            return
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
