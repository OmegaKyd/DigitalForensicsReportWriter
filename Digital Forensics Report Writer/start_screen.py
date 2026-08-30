import sys
sys.dont_write_bytecode = True

import tkinter as tk
from tkinter import ttk, messagebox
import ctypes

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


class StartScreen(tk.Tk):
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
        ttk.Button(footer, text="Exit", command=self.destroy, width=12).pack(side="right")

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

    def launch_mobile_portable_case(self):
        self.withdraw()
        app = MobilePortableCase(self)
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_app_close(app))

    def launch_mobile_full_exam(self):
        self.withdraw()
        app = MobileFullExam(self)
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_app_close(app))

    def launch_pc_portable_case(self):
        self.withdraw()
        app = PCPortableCase(self)
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_app_close(app))

    def launch_pc_full_exam(self):
        self.withdraw()
        app = PCFullExam(self)
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_app_close(app))

    def launch_warrant_data_returns(self):
        self.withdraw()
        app = WarrantDataReturns(self)
        app.protocol("WM_DELETE_WINDOW", lambda: self.on_app_close(app))

    def on_close(self):
        self.destroy()
        sys.exit()

    def on_app_close(self, app):
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

# Ω Digital Forensics Report Writer Ω (ver. 1.0.1) © 2026 #
