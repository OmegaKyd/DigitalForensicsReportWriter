import sys
sys.dont_write_bytecode = True

import tkinter as tk
from tkinter import ttk, messagebox
import ctypes
from tkinterdnd2 import TkinterDnD

from mobile_portable_case import MobilePortableCase
from mobile_full_exam import MobileFullExam
from pc_full_exam import PCFullExam
from pc_portable_case import PCPortableCase
from sw_data_review import WarrantDataReturns
from ui_theme import APP_NAME, apply_theme, add_header_bar, add_action_bar, close_and_return
from app_menu import attach_app_menu


def _hide_console():
    try:
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    except Exception:
        pass


class StartScreen(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.title(f"Ω {APP_NAME} Ω")
        apply_theme(self)
        attach_app_menu(self, is_start=True)

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        width = 900
        height = min(640, max(560, screen_height - 80))
        x = (screen_width - width) // 2
        y = 40
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(780, 560)

        add_header_bar(self, APP_NAME, "Case reports for mobile, computer, and warrant returns")
        footer = add_action_bar(self)
        ttk.Button(footer, text="Exit", command=self.on_close, width=12).pack(side="right")
        self.open_app = None

        self.main_frame = ttk.Frame(self, padding="18")
        self.main_frame.pack(expand=True, fill="both")

        ttk.Label(self.main_frame, text="Select a report type to begin").pack(pady=(0, 14))

        columns_frame = ttk.Frame(self.main_frame)
        columns_frame.pack(expand=True, fill="both")
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)

        frame_width = 380

        mobile_frame = ttk.Frame(columns_frame, style="CategoryFrame.TFrame", padding=16, width=frame_width)
        mobile_frame.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        ttk.Label(mobile_frame, text="Mobile Device Reports", style="CategoryHeader.TLabel").pack(pady=(0, 10), anchor="center")
        ttk.Button(mobile_frame, text="Mobile — Portable Case", command=self.launch_mobile_portable_case, width=32).pack(pady=5)
        ttk.Button(mobile_frame, text="Mobile — Full Exam", command=self.launch_mobile_full_exam, width=32).pack(pady=5)
        ttk.Label(
            mobile_frame,
            text="For Cellebrite UFD/Summary/Quick View and GrayKey PDF extractions.",
            wraplength=300,
            style="Hint.TLabel",
            justify="center",
        ).pack(pady=(8, 4))

        warrant_frame = ttk.Frame(columns_frame, style="CategoryFrame.TFrame", padding=16, width=frame_width)
        warrant_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)

        ttk.Label(warrant_frame, text="Search Warrant Data Returns", style="CategoryHeader.TLabel").pack(pady=(0, 10), anchor="center")
        ttk.Button(warrant_frame, text="Warrant Data Returns Report", command=self.launch_warrant_data_returns, width=32).pack(pady=5)
        ttk.Label(
            warrant_frame,
            text="For reports from warrant returns, subpoenas, or service provider data.",
            wraplength=300,
            style="Hint.TLabel",
            justify="center",
        ).pack(pady=(8, 4))

        pc_frame = ttk.Frame(columns_frame, style="CategoryFrame.TFrame", padding=16, width=frame_width)
        pc_frame.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)

        ttk.Label(pc_frame, text="Computer Reports", style="CategoryHeader.TLabel").pack(pady=(0, 10), anchor="center")
        ttk.Button(pc_frame, text="PC — Portable Case", command=self.launch_pc_portable_case, width=32).pack(pady=5)
        ttk.Button(pc_frame, text="PC — Full Exam", command=self.launch_pc_full_exam, width=32).pack(pady=5)
        ttk.Label(
            pc_frame,
            text="Use the Image Acquisition Log from TX1, FTK, X-Ways, or Cellebrite Digital Collector.",
            wraplength=300,
            style="Hint.TLabel",
            justify="center",
        ).pack(pady=(8, 4))

        columns_frame.rowconfigure(0, weight=1)
        columns_frame.rowconfigure(1, weight=1)

        self.after_idle(self._ensure_templates)

    def _ensure_templates(self):
        from report_common import ensure_user_templates
        ensure_user_templates(self)

    def _close_open_app(self):
        app = getattr(self, "open_app", None)
        self.open_app = None
        if app is None:
            return
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
        try:
            app.grab_release()
        except Exception:
            pass
        try:
            app.protocol("WM_DELETE_WINDOW", lambda: None)
        except Exception:
            pass
        try:
            if app.winfo_exists():
                app.withdraw()
                app.destroy()
        except Exception:
            try:
                if app.winfo_exists():
                    app.destroy()
            except Exception:
                pass

    def _launch(self, factory):
        self._close_open_app()
        self.withdraw()
        app = factory(self)
        self.open_app = app
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_app_close(app))
        draft = getattr(self, "_pending_draft", None)
        self._pending_draft = None
        if draft:
            def _apply():
                from drafts_manager import apply_draft
                try:
                    apply_draft(app, draft)
                except Exception as exc:
                    messagebox.showerror("Open Unfinished Reports", str(exc), parent=app)
            app.after_idle(_apply)

    def launch_mobile_portable_case(self):
        self._launch(MobilePortableCase)

    def launch_mobile_full_exam(self):
        self._launch(MobileFullExam)

    def launch_pc_portable_case(self):
        self._launch(PCPortableCase)

    def launch_pc_full_exam(self):
        self._launch(PCFullExam)

    def launch_warrant_data_returns(self):
        self._launch(WarrantDataReturns)

    def on_close(self):
        self._close_open_app()
        try:
            self.destroy()
        except Exception:
            pass
        sys.exit()

    def on_app_close(self, app):
        if getattr(self, "open_app", None) is app:
            self.open_app = None
        close_and_return(app)


if __name__ == "__main__":
    _hide_console()
    try:
        root = StartScreen()
        root.mainloop()
    except Exception as e:
        try:
            messagebox.showerror("Error", f"An unexpected error occurred: {str(e)}")
        except Exception:
            print(f"An unexpected error occurred: {e}")
        sys.exit(1)

# Ω Digital Forensics Report Writer Ω (ver. 1.0.5) © 2026 #
