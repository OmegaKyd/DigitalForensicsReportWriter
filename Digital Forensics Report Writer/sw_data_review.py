import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from docx import Document
from docx.shared import Pt
import os
import re
from tkinterdnd2 import DND_FILES, TkinterDnD
from settings_manager import SettingsManager
from ui_theme import apply_theme, add_header_bar, add_action_bar, pack_right_actions, size_window, parse_mdy_date, add_date_entry, COLORS, close_and_return, DndToplevel
from app_menu import attach_app_menu
from paragraphs_manager import fill_paragraph, load_paragraphs
from report_common import (
    current_dfr_prefix,
    is_complete_dfr_number,
    add_template_picker,
    sync_template_choice,
    load_request_titles,
    setup_title_combobox,
    remember_titles_from_form,
    setup_agency_combobox,
    remember_agencies_from_form,
    title_agency_words,
    saved_examiner_agency,
    bind_prefix_typeahead,
    refresh_request_title_values,
    apply_warrant_suggested_filename,
)
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from lxml import etree

#ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

class WarrantDataReturns(DndToplevel):
    def __init__(self, master=None):
        if master is None:
            master = TkinterDnD.Tk()
            master.withdraw()
            owns_root = True
        else:
            owns_root = False
        super().__init__(master)
        self._owns_hidden_root = owns_root
        
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.master = master
        apply_theme(self)
        attach_app_menu(self)
        add_header_bar(self, "Warrant Data Returns", "Warrant, subpoena, and service-provider reports")
        self.action_bar = add_action_bar(self)
        self.generate_button = ttk.Button(self.action_bar, text="Generate Report", command=self.generate_report)
        self.exit_button = ttk.Button(self.action_bar, text="Exit", command=self.on_closing)
        pack_right_actions(self.generate_button, self.exit_button)
        
        self.settings_manager = SettingsManager()
        self.current_settings = self.settings_manager.load_settings()
        self._save_timer = None
        
        self.paragraphs = {}
        self._paragraph_kind = "warrant"
        self.initialize_warrant_paragraphs()
        
        self.title("Ω Digital Forensics Report Writer - Warrant Data Returns Ω")
        
        size_window(self, 1120, 720, min_width=980, min_height=660)
        
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        self.columns_frame = ttk.Frame(self.main_frame)
        self.columns_frame.pack(fill=tk.X, expand=False, padx=8, pady=8)
        
        # Left Column 
        self.left_frame = ttk.Frame(self.columns_frame, padding=0)
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        
        self.scrollable_frame = ttk.Frame(self.left_frame, padding=(2, 2, 2, 2))
        self.scrollable_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Middle Column
        self.middle_frame = ttk.Frame(self.columns_frame, padding=0)
        self.middle_frame.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        
        # Right Column
        self.right_frame = ttk.Frame(self.columns_frame, padding=0)
        self.right_frame.grid(row=0, column=2, sticky="nsew", padx=0, pady=0)
        
        self.columns_frame.columnconfigure(0, weight=1, minsize=320)
        self.columns_frame.columnconfigure(1, weight=1, minsize=320)
        self.columns_frame.columnconfigure(2, weight=1, minsize=320)
        self.columns_frame.rowconfigure(0, weight=0)
        
        self.create_widgets()

    def create_widgets(self):
        self.create_left_column_widgets()
        self.create_middle_column_widgets()
        self.create_right_column_widgets()

# LEFT PANE
    def create_left_column_widgets(self):
        self.create_role_selection_frame()
        self.create_request_information_frame()
        self.create_warrant_return_dates_frame()
        self.create_examiner_information_frame()

    def create_role_selection_frame(self):
        role_frame = ttk.LabelFrame(self.scrollable_frame, text="Examination Type", padding=(4, 2))
        role_frame.pack(fill=tk.X, padx=2, pady=2)
        ttk.Label(role_frame, text="Select Role:").grid(row=0, column=0, sticky="w", pady=2)
        self.role_type = ttk.Combobox(
            role_frame,
            values=["Agency Assist", "Case Agent"],
            state="readonly",
        )
        self.role_type.grid(row=0, column=1, sticky="ew", pady=2)
        self.role_type.set("Agency Assist")
        self.role_type.bind("<<ComboboxSelected>>", self.toggle_request_content)
        role_frame.columnconfigure(1, weight=1)

    def toggle_request_content(self, event=None):
        if not hasattr(self, "agency_assist_fields"):
            return
        if self.role_type.get() == "Case Agent":
            self.agency_assist_fields.grid_remove()
        else:
            self.agency_assist_fields.grid()

    def create_warrant_return_dates_frame(self):
        warrant_frame = ttk.LabelFrame(self.scrollable_frame, text="Warrant Return Information", padding=6)
        warrant_frame.pack(fill=tk.X, padx=5, pady=4)

        ttk.Label(warrant_frame, text="Warrant Service Date:").grid(row=0, column=0, sticky="w", pady=2)
        self.warrant_service_date = add_date_entry(warrant_frame, row=0, column=1)

        ttk.Label(warrant_frame, text="Data Return Date:").grid(row=1, column=0, sticky="w", pady=2)
        self.data_return_date = add_date_entry(warrant_frame, row=1, column=1)

        self.time_frame_label = ttk.Label(warrant_frame, text="Time Limit Limited?:")
        self.time_frame_label.grid(row=2, column=0, sticky="w", pady=2)

        self.time_frame_checkbox_frame = ttk.Frame(warrant_frame)
        self.time_frame_checkbox_frame.grid(row=2, column=1, sticky="ew", pady=2)

        self.time_frame_limited_var = tk.IntVar(warrant_frame)
        self.time_frame_checkbox = ttk.Checkbutton(
            self.time_frame_checkbox_frame,
            text="Yes",
            variable=self.time_frame_limited_var,
            onvalue=1, offvalue=0,
            command=self.toggle_time_frame_fields
        )
        self.time_frame_checkbox.grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.time_frame_dates_frame = ttk.Frame(warrant_frame)
        self.time_frame_dates_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=2)
        self.time_frame_dates_frame.grid_remove()

        ttk.Label(self.time_frame_dates_frame, text="Time Frame Start Date:").grid(row=0, column=0, sticky="w", pady=2)
        self.time_frame_start_date = add_date_entry(self.time_frame_dates_frame, row=0, column=1)

        ttk.Label(self.time_frame_dates_frame, text="Time Frame End Date:").grid(row=1, column=0, sticky="w", pady=2)
        self.time_frame_end_date = add_date_entry(self.time_frame_dates_frame, row=1, column=1)

        warrant_frame.columnconfigure(1, weight=1)

    def create_request_information_frame(self):
        request_frame = ttk.LabelFrame(self.scrollable_frame, text="Request Information", padding=(4, 2))
        request_frame.pack(fill=tk.X, padx=2, pady=2)

        self.agency_assist_fields = ttk.Frame(request_frame)
        self.agency_assist_fields.grid(row=0, column=0, columnspan=2, sticky="ew")

        ttk.Label(self.agency_assist_fields, text="Request Date:").grid(row=0, column=0, sticky="w", pady=2)
        self.request_date = add_date_entry(self.agency_assist_fields, row=0, column=1)

        ttk.Label(self.agency_assist_fields, text="Requesting Agency:").grid(row=1, column=0, sticky="w", pady=2)
        self.requesting_agency = ttk.Combobox(self.agency_assist_fields)
        self.requesting_agency.grid(row=1, column=1, sticky="ew", pady=2)
        setup_agency_combobox(self.requesting_agency)

        ttk.Label(self.agency_assist_fields, text="Requesting Officer Title:").grid(row=2, column=0, sticky="w", pady=2)
        self.requesting_officer_title = ttk.Combobox(self.agency_assist_fields, values=load_request_titles())
        self.requesting_officer_title.grid(row=2, column=1, sticky="ew", pady=2)
        self.requesting_officer_title.set("")
        bind_prefix_typeahead(self.requesting_officer_title)

        self.requesting_officer_title_other = ttk.Entry(self.agency_assist_fields)
        self.requesting_officer_title_other.grid(row=3, column=1, sticky="ew", pady=2)
        self.requesting_officer_title_other.grid_remove()

        ttk.Label(self.agency_assist_fields, text="Requesting Officer:").grid(row=4, column=0, sticky="w", pady=2)
        self.requesting_officer = ttk.Entry(self.agency_assist_fields)
        self.requesting_officer.grid(row=4, column=1, sticky="ew", pady=2)
        self.agency_assist_fields.columnconfigure(1, weight=1)

        ttk.Label(request_frame, text="Primary Case Offense:").grid(row=1, column=0, sticky="w", pady=2)
        self.offense_var = tk.StringVar()
        self.primary_case_offense = ttk.Entry(request_frame, textvariable=self.offense_var)
        self.primary_case_offense.grid(row=1, column=1, sticky="ew", pady=2)
        self.offense_var.trace_add("write", lambda *args: self.format_offense_lowercase())

        request_frame.columnconfigure(1, weight=1)

    def format_offense_lowercase(self):
        current = self.offense_var.get()
        if current:
            formatted = ''.join(c.lower() if c.isalpha() else c for c in current)
            if formatted != current:
                self.offense_var.set(formatted)

    def toggle_requesting_officer_title_other(self, event=None):
        if self.requesting_officer_title.get() == "Other":
            self.requesting_officer_title_other.grid()
        else:
            self.requesting_officer_title_other.grid_remove()
            self.requesting_officer_title_other.delete(0, tk.END)

    def create_examiner_information_frame(self):
        examiner_frame = ttk.LabelFrame(self.scrollable_frame, text="Examiner Information", padding=(4, 2))
        examiner_frame.pack(fill=tk.X, padx=2, pady=2)

        ttk.Label(examiner_frame, text="Examiner Agency:").grid(row=0, column=0, sticky="w", pady=2)
        self.examiner_agency = ttk.Combobox(examiner_frame)
        self.examiner_agency.grid(row=0, column=1, sticky="ew", pady=2)
        setup_agency_combobox(
            self.examiner_agency,
            saved=saved_examiner_agency(self.current_settings),
            on_change=self.auto_save_settings,
        )

        self.examiner_agency_other = ttk.Entry(examiner_frame)
        self.examiner_agency_other.grid(row=1, column=1, sticky="ew", pady=2)
        self.examiner_agency_other.grid_remove()

        ttk.Label(examiner_frame, text="Examiner Title:").grid(row=2, column=0, sticky="w", pady=2)
        self.examiner_title = ttk.Combobox(examiner_frame, values=load_request_titles())
        self.examiner_title.grid(row=2, column=1, sticky="ew", pady=2)
        setup_title_combobox(
            self.examiner_title,
            saved=self.current_settings.get("examiner_title", ""),
            on_change=self.auto_save_settings,
        )

        self.examiner_title_other = ttk.Entry(examiner_frame)
        self.examiner_title_other.grid(row=3, column=1, sticky="ew", pady=2)
        self.examiner_title_other.grid_remove()

        ttk.Label(examiner_frame, text="Examiner Name:").grid(row=4, column=0, sticky="w", pady=2)
        self.examiner_name = ttk.Entry(examiner_frame)
        saved_name = self.current_settings.get("examiner_name", "")
        self.examiner_name.insert(0, saved_name)
        self.examiner_name.grid(row=4, column=1, sticky="ew", pady=2)

        ttk.Label(examiner_frame, text="Case Number:").grid(row=5, column=0, sticky="w", pady=2)
        self.case_var = tk.StringVar() 
        self.case_number = ttk.Entry(examiner_frame, textvariable=self.case_var)
        self.case_number.grid(row=5, column=1, sticky="ew", pady=2)
        self.case_var.trace_add("write", lambda *args: self.format_case_number_upper())

        ttk.Label(examiner_frame, text="DFR Number:").grid(row=6, column=0, sticky="w", pady=2)
        self.dfr_number = ttk.Entry(examiner_frame)

        self.dfr_number.insert(0, current_dfr_prefix())
        self.dfr_number.grid(row=6, column=1, sticky="ew", pady=2)

        examiner_frame.columnconfigure(1, weight=1)

    def format_case_number_upper(self):
        current = self.case_var.get()
        if current:
            formatted = ''.join(c.upper() if c.isalpha() else c for c in current)
            if formatted != current:
                self.case_var.set(formatted)

    def toggle_examiner_agency_other(self, event=None):
        return

    def toggle_examiner_title_other(self, event=None):
        if self.examiner_title.get() == "Other":
            self.examiner_title_other.grid()
        else:
            self.examiner_title_other.grid_remove()
            self.examiner_title_other.delete(0, tk.END)

    def toggle_time_frame_fields(self):
        if self.time_frame_limited_var.get() == 1:
            self.time_frame_dates_frame.grid()
        else:
            self.time_frame_dates_frame.grid_remove()
            if hasattr(self, 'time_frame_start_date'):
                self.time_frame_start_date.delete(0, tk.END)
            if hasattr(self, 'time_frame_end_date'):
                self.time_frame_end_date.delete(0, tk.END)

# MIDDLE PANE
    def create_middle_column_widgets(self):
        self.create_account_information_frame()
        self.create_forensic_software_frame()

    def create_account_information_frame(self):
        account_frame = ttk.LabelFrame(self.middle_frame, text="Account Information", padding=6)
        account_frame.pack(fill=tk.X, padx=5, pady=4, expand=False)   # ← no expand

        ttk.Label(account_frame, text="Service Provider:").grid(row=0, column=0, sticky="w", pady=2)
        self.service_provider = ttk.Entry(account_frame)
        self.service_provider.grid(row=0, column=1, sticky="ew", pady=2)

        ttk.Label(account_frame, text="Account Identifier:").grid(row=1, column=0, sticky="w", pady=2)
        self.account_identifier = ttk.Entry(account_frame)
        self.account_identifier.grid(row=1, column=1, sticky="ew", pady=2)

        ttk.Label(account_frame, text="GB of Data:").grid(row=2, column=0, sticky="w", pady=2)
        self.account_data_size = ttk.Entry(account_frame)
        self.account_data_size.grid(row=2, column=1, sticky="ew", pady=2)

        ttk.Label(account_frame, text="Account Owner:").grid(row=3, column=0, sticky="w", pady=2)
        self.account_owner = ttk.Entry(account_frame)
        self.account_owner.grid(row=3, column=1, sticky="ew", pady=2)

        account_frame.columnconfigure(1, weight=1)

    def create_forensic_software_frame(self):
        forensic_frame = ttk.LabelFrame(self.middle_frame, text="Forensic Processing Software", padding=6)
        forensic_frame.pack(fill=tk.X, padx=5, pady=4, expand=False)

        self.cb_cellebrite_var = tk.IntVar(forensic_frame)
        self.cb_axiom_var      = tk.IntVar(forensic_frame)
        self.cb_griffeye_var   = tk.IntVar(forensic_frame)
        self.cb_manual_var     = tk.IntVar(forensic_frame)

        ttk.Checkbutton(
            forensic_frame, text="Cellebrite", variable=self.cb_cellebrite_var,
            onvalue=1, offvalue=0, command=self.on_processing_software_toggled
        ).grid(row=0, column=0, sticky="w", padx=(0, 0), pady=4)

        ttk.Checkbutton(
            forensic_frame, text="AXIOM", variable=self.cb_axiom_var,
            onvalue=1, offvalue=0, command=self.on_processing_software_toggled
        ).grid(row=0, column=1, sticky="w", padx=(0, 0), pady=4)

        ttk.Checkbutton(
            forensic_frame, text="Griffeye", variable=self.cb_griffeye_var,
            onvalue=1, offvalue=0, command=self.on_processing_software_toggled
        ).grid(row=0, column=2, sticky="w", padx=(0, 0), pady=4)

        ttk.Checkbutton(
            forensic_frame, text="Manual Exam Only", variable=self.cb_manual_var,
            onvalue=1, offvalue=0, command=self.on_manual_exam_toggled
        ).grid(row=0, column=3, sticky="w", padx=(0, 0), pady=4)

        for col in range(4):
            forensic_frame.columnconfigure(col, weight=1)

    def get_selected_forensic_software(self):
        selected = []
        if hasattr(self, 'cb_cellebrite_var') and self.cb_cellebrite_var.get() == 1:
            selected.append("Cellebrite")
        if hasattr(self, 'cb_axiom_var') and self.cb_axiom_var.get() == 1:
            selected.append("AXIOM")
        if hasattr(self, 'cb_griffeye_var') and self.cb_griffeye_var.get() == 1:
            selected.append("Griffeye")
        if hasattr(self, 'cb_manual_var') and self.cb_manual_var.get() == 1:
            selected.append("Manual Exam Only")
        return selected

    def on_manual_exam_toggled(self):
        if self.cb_manual_var.get() != 1:
            return
        self.cb_cellebrite_var.set(0)
        self.cb_axiom_var.set(0)
        self.cb_griffeye_var.set(0)

    def on_processing_software_toggled(self):
        if (
            self.cb_cellebrite_var.get() == 1
            or self.cb_axiom_var.get() == 1
            or self.cb_griffeye_var.get() == 1
        ):
            self.cb_manual_var.set(0)

    def get_processing_software(self):
        """Cellebrite / AXIOM / Griffeye only. Manual Exam Only is not processing software."""
        return [name for name in self.get_selected_forensic_software() if name != "Manual Exam Only"]

    def format_software_list(self, names):
        items = [str(name).strip() for name in (names or []) if str(name).strip()]
        if not items:
            return ""
        if len(items) == 1:
            return items[0]
        if len(items) == 2:
            return f"{items[0]} and {items[1]}"
        return ", ".join(items[:-1]) + ", and " + items[-1]

    def report_article_for(self, software_name):
        text = (software_name or "").strip()
        if not text:
            return "a"
        return "an" if text[0].lower() in "aeiou" else "a"

    def ask_digital_report_software(self, tools):
        """Ask which checked tools produced a Digital Report. None means cancel."""
        result = {"tools": None}
        win = tk.Toplevel(self)
        win.title("Digital Report Software")
        win.configure(bg=COLORS["bg"])
        win.transient(self)
        win.resizable(False, False)

        ttk.Label(
            win,
            text="More than one forensic tool is selected. Which of the selected software did you use to create a Digital Report?",
            wraplength=420,
        ).pack(anchor="w", padx=16, pady=(14, 8))
        ttk.Label(
            win,
            text="Check all that apply.",
            wraplength=420,
        ).pack(anchor="w", padx=16, pady=(0, 8))

        vars_by_name = {}
        for name in tools:
            var = tk.IntVar(win, value=0)
            vars_by_name[name] = var
            ttk.Checkbutton(win, text=name, variable=var, onvalue=1, offvalue=0).pack(anchor="w", padx=24, pady=2)

        def accept():
            chosen = [name for name, var in vars_by_name.items() if var.get() == 1]
            if not chosen:
                messagebox.showinfo(
                    "Digital Report Software",
                    "Select at least one tool, or click Cancel.",
                    parent=win,
                )
                return
            result["tools"] = chosen
            win.destroy()

        def cancel():
            result["tools"] = None
            win.destroy()

        buttons = ttk.Frame(win)
        buttons.pack(fill="x", padx=16, pady=14)
        ttk.Button(buttons, text="Use these tools", style="Accent.TButton", command=accept).pack(side="right")
        ttk.Button(buttons, text="Cancel", command=cancel).pack(side="right", padx=(0, 8))

        win.protocol("WM_DELETE_WINDOW", cancel)
        win.update_idletasks()
        width = max(460, win.winfo_reqwidth())
        height = win.winfo_reqheight()
        px = self.winfo_rootx() + max(0, (self.winfo_width() - width) // 2)
        py = self.winfo_rooty() + max(0, (self.winfo_height() - height) // 3)
        win.geometry(f"{width}x{height}+{px}+{py}")
        win.grab_set()
        win.focus_set()
        self.wait_window(win)
        return result["tools"]

    def create_output_file_frame(self):
        output_frame = ttk.LabelFrame(self.right_frame, text="Output File", padding="10")
        output_frame.pack(fill=tk.X, padx=5, pady=5)
        content = ttk.Frame(output_frame)
        content.pack(fill=tk.X, expand=True)

        # Output File Name
        ttk.Label(content, text="Output File Name:").grid(row=0, column=0, sticky="w", pady=2)
        self.output_filename = ttk.Entry(content)
        self.output_filename.grid(row=1, column=0, columnspan=2, sticky="ew", pady=2)

        # Hint text
        hint_style = ttk.Style()
        hint_style.configure("Hint.TLabel", font=('Arial', 8))
        hint_label = ttk.Label(content,
                               text="*Default File Name \"(DFR #) - (Provider) (AccountID) Return\"",
                               style="Hint.TLabel", wraplength=350)
        hint_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 6))

        # Save Location label
        ttk.Label(content, text="Save Location:").grid(row=3, column=0, sticky="w", pady=2)

        # Save Location entry (full width, no sub-frame)
        self.save_location = ttk.Entry(content)
        self.save_location.grid(row=4, column=0, columnspan=2, sticky="ew", pady=2)
        self.save_location.insert(0, os.path.join(os.path.expanduser("~"), "Desktop"))

        # Browse button on its own row below the entry
        ttk.Button(content, text="Browse", command=self.browse_save_location).grid(row=5, column=0, columnspan=1, pady=4, sticky="ew")

        content.columnconfigure(1, weight=1)
        for widget in (self.dfr_number, self.service_provider, self.account_identifier):
            widget.bind("<KeyRelease>", lambda event: apply_warrant_suggested_filename(self), add="+")
        apply_warrant_suggested_filename(self)

    def create_template_file_frame(self):
        template_frame = ttk.LabelFrame(self.right_frame, text="Template File", padding=6)
        template_frame.pack(fill=tk.X, expand=False, padx=5, pady=4, anchor="n")
        add_template_picker(self, template_frame, preferred="DFR SW Return (2026).docx", keywords=("sw", "warrant", "return"))
         
    def on_template_drop(self, event):
        """Handle drag-and-drop of DFR template file."""
        try:
            file_path = self.tk.splitlist(event.data)[0]

            if not file_path.lower().endswith('.docx'):
                messagebox.showerror("Invalid File", "Please drop a .docx file.")
                return

            self.template_file = file_path
            sync_template_choice(self, file_path)
            self.template_status_label.configure(text=f"Selected: {os.path.basename(file_path)}")

        except Exception as e:
            messagebox.showerror("Drop Error", f"Failed to process drop:\n{str(e)}")

    def browse_template(self, event=None):
        """Browse for DFR template file."""
        path = filedialog.askopenfilename(
            title="Select DFR Template",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")]
        )
        if path:
            self.template_file = path
            sync_template_choice(self, path)
            self.template_status_label.configure(text=f"Selected: {os.path.basename(path)}")

    def create_right_column_widgets(self):
        self.create_template_file_frame()
        self.create_output_file_frame()

    # ================== HELPERS ==================
    def toggle_title_entry(self, event=None):
        self.toggle_requesting_officer_title_other(event)

    def toggle_examiner_title_entry(self, event=None):
        self.toggle_examiner_title_other(event)

    def on_agency_type_changed(self, event=None):
        self.toggle_examiner_agency_other(event)
        self.auto_save_settings()

    def toggle_agency_entry(self, event=None):
        self.toggle_examiner_agency_other(event)

    def auto_save_settings(self, event=None):
        if hasattr(self, '_save_timer') and self._save_timer:
            self.after_cancel(self._save_timer)
        self._save_timer = self.after(500, self._perform_auto_save)

    def _perform_auto_save(self):
        try:
            # Extract ONLY the prefix from the current DFR field
            current_dfr = self.dfr_number.get().strip()
            prefix = current_dfr

            # If there's a hyphen, take everything up to and including the hyphen
            if '-' in current_dfr:
                prefix = current_dfr.rsplit('-', 1)[0] + '-'

            # Clean and validate prefix
            prefix = prefix.strip()
            if not prefix.endswith('-'):
                prefix += '-'

            # Ensure it's a valid DFR prefix (DFR + year + -)
            if not (prefix.startswith('DFR') and len(prefix) >= 8 and prefix[3:7].isdigit()):
                current_year = datetime.now().year
                prefix = f"DFR{current_year}-"

            settings_to_save = {
                "examiner_agency_type": self.examiner_agency.get(),
                "examiner_agency_custom": "",
                "examiner_title": self.examiner_title.get(),
                "examiner_title_custom": self.examiner_title_other.get() if self.examiner_title.get() == "Other" else "",
                "examiner_name": self.examiner_name.get(),
                "dfr_number_prefix": prefix,  # always just the prefix (e.g. "DFR2026-")
            }
            self.settings_manager.save_settings(settings_to_save)
            self.current_settings.update(settings_to_save)
            remember_agencies_from_form(self)
        except Exception as e:
            print(f"Error auto-saving: {e}")

    def get_dfr_prefix(self):
        dfr = self.dfr_number.get() if hasattr(self, "dfr_number") else ""
        match = re.match(r'(DFR\d{4}-)', dfr)
        return match.group(1) if match else f"DFR{datetime.now().year}-"

    def setup_auto_complete_dropdown(self, combobox):
        combobox.bind('<KeyRelease>', lambda event, cb=combobox: self.auto_complete(event, cb))

    def auto_complete(self, event, combobox):
        if event.keysym in ('Up', 'Down', 'Left', 'Right', 'Return', 'Tab', 'BackSpace', 'Delete'):
            return
        value = combobox.get().lower()
        if not value:
            return
        for option in combobox['values']:
            if option.lower().startswith(value):
                combobox.set(option)
                combobox.icursor(len(option))
                combobox.selection_range(len(value), len(option))
                break

    def browse_save_location(self):
        current = self.save_location.get().strip() or os.path.join(os.path.expanduser("~"), "Desktop")
        selected = filedialog.askdirectory(title="Select Save Location", initialdir=current)
        if selected:
            self.save_location.delete(0, tk.END)
            self.save_location.insert(0, selected)

###UPDATE PARAGRAPHS###

    def initialize_warrant_paragraphs(self):
        self.paragraphs = load_paragraphs("warrant")

    def generate_paragraphs(self, data, new_doc):
        def add(text):
            p = new_doc.add_paragraph(fill_paragraph(text, data))
            for run in p.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(11)

        role = data.get("Role_Type") or (self.role_type.get() if hasattr(self, "role_type") else "Agency Assist")
        if role == "Case Agent":
            limited = bool(data.get("Time_Frame_Start") and data.get("Time_Frame_End"))
            if limited:
                add(self.paragraphs["one_ca_limited"])
            else:
                add(self.paragraphs["one_ca"])
            add(self.paragraphs["two_ca"])
        else:
            add(self.paragraphs["one_aa"])
            add(self.paragraphs["two_aa"])

        processing = [name for name in (data.get("Processing_Software_List") or []) if name]
        if processing:
            add(self.paragraphs["three_ca"])
            report_tools = data.get("Report_Software_List") or processing
            if len(report_tools) > 1:
                add(self.paragraphs["four_ca_multi"])
            else:
                add(self.paragraphs["four_ca"])
        else:
            add(self.paragraphs["three_b"])

###PARAGRAPHS###

    def validate_fields(self):
        missing = []

        # Examiner Agency
        examiner_agency = self.examiner_agency.get().strip()
        if not examiner_agency:
            missing.append("Examiner Agency")

        # Examiner Title
        examiner_title = self.examiner_title.get().strip()
        if not examiner_title:
            missing.append("Examiner Title")

        # Examiner Name
        if not self.examiner_name.get().strip():
            missing.append("Examiner Name")

        # Case Number
        if not self.case_number.get().strip():
            missing.append("Case Number")

        # DFR Number
        if not is_complete_dfr_number(self.dfr_number.get()):
            missing.append("DFR Number")

        role = self.role_type.get() if hasattr(self, "role_type") else "Agency Assist"
        if role != "Case Agent":
            if not self.request_date.get().strip():
                missing.append("Request Date")
            if not self.requesting_agency.get().strip():
                missing.append("Requesting Agency")
            if not self.requesting_officer_title.get().strip():
                missing.append("Requesting Officer Title")
            if not self.requesting_officer.get().strip():
                missing.append("Requesting Officer")

        # Primary Case Offense
        if not self.primary_case_offense.get().strip():
            missing.append("Primary Case Offense")

        # Warrant Service Date
        if not self.warrant_service_date.get().strip():
            missing.append("Warrant Service Date")

        # Data Return Date
        if not self.data_return_date.get().strip():
            missing.append("Data Return Date")

        if hasattr(self, "time_frame_limited_var") and self.time_frame_limited_var.get() == 1:
            if not self.time_frame_start_date.get().strip():
                missing.append("Time Frame Start Date")
            if not self.time_frame_end_date.get().strip():
                missing.append("Time Frame End Date")

        if not self.service_provider.get().strip():
            missing.append("Service Provider")
        if not self.account_identifier.get().strip():
            missing.append("Account Identifier")
        if not self.account_data_size.get().strip():
            missing.append("GB of Data")
        if not self.account_owner.get().strip():
            missing.append("Account Owner")

        if not self.get_selected_forensic_software():
            missing.append("Forensic Processing Software or Manual Exam Only")

        # Template File - use the correct attribute name
        template_path = getattr(self, 'template_file', None)
        if not template_path or not os.path.exists(template_path):
            missing.append(f"DFR Template File (current path: {template_path or 'None'})")

        return missing

    def generate_report(self):
        missing = self.validate_fields()
        if missing:
            messagebox.showerror("Missing Fields", f"Please fill in the following:\n\n" + "\n".join(f"• {f}" for f in missing))
            return

        if not hasattr(self, 'template_file') or not self.template_file:
            messagebox.showerror("Error", "No template file selected.")
            return

        examiner_title = self.format_title(self.examiner_title.get())
        examiner_name = self.examiner_name.get().strip().title()
        examiner_value = f"{examiner_title} {examiner_name}".strip()
        examiner_agency, examiner_agency_abbr = self.format_agency(
            self.examiner_agency.get().strip(), return_abbreviation=True
        )
        role = self.role_type.get() if hasattr(self, "role_type") else "Agency Assist"
        if role == "Case Agent":
            request_title = examiner_title
            request_officer = examiner_name
            request_agency = examiner_agency
            request_agency_abbr = examiner_agency_abbr
            request_officer_full = examiner_value
        else:
            request_title = self.format_title(self.requesting_officer_title.get())
            request_officer = self.requesting_officer.get().strip().title()
            request_agency, request_agency_abbr = self.format_agency(
                self.requesting_agency.get().strip(), return_abbreviation=True
            )
            request_officer_full = f"{request_title} {request_officer}".strip()

        processing_tools = self.get_processing_software()
        report_tools = list(processing_tools)
        if len(processing_tools) > 1:
            chosen = self.ask_digital_report_software(processing_tools)
            if chosen is None:
                return
            report_tools = chosen

        forensic_software_text = self.format_software_list(processing_tools)
        report_software_text = self.format_software_list(report_tools)
        report_article = self.report_article_for(report_tools[0] if report_tools else "")

        # Gather data using current widgets
        data = {
            'PY_DFR': self.dfr_number.get().strip().upper(),
            'PY_CASENUMBER': self.case_number.get().strip().upper(),
            'PY_EXAMINER': examiner_value,
            'PY_REQAGENCY': request_agency,
            'PY_REQOFF': request_officer_full,
            'PY_SERVEDATE': self.parse_request_date(self.warrant_service_date.get().strip()),
            'PY_RETURNDATE': self.parse_request_date(self.data_return_date.get().strip()),
            'PY_PROVIDER': self.service_provider.get().strip(),
            'PY_ACCOUNTID': self.account_identifier.get().strip(),
            'PY_DATASIZE': self.account_data_size.get().strip(),
            'PY_OWNER': self.account_owner.get().strip().title(),
            'Account_Owner': self.account_owner.get().strip().title(),
            'Account_Identifier': self.account_identifier.get().strip(),
            'Service_Provider': self.service_provider.get().strip(),
            'Data_Size': self.account_data_size.get().strip(),
            'Warrant_Service_Date': self.parse_request_date(self.warrant_service_date.get().strip()),
            'Data_Return_Date': self.parse_request_date(self.data_return_date.get().strip()),
            'DFR_Num': self.dfr_number.get().strip().upper(),
            'Case_Number': self.case_number.get().strip().upper(),
            'Examiner_Title': examiner_title,
            'Examiner_Name': examiner_name,
            'Examiner_Agency': examiner_agency,
            'Examiner_Agency_Abbr': examiner_agency_abbr,
            'Request_Agency': request_agency,
            'Request_Agency_Abbr': request_agency_abbr,
            'Request_Title': request_title,
            'Request_Officer': request_officer,
            'Request_Officer_Full': request_officer_full,
            'Request_Date': self.parse_request_date(self.request_date.get().strip()) if hasattr(self, "request_date") else "",
            'Request_Case': self.primary_case_offense.get().strip(),
            'Role_Type': role,
            'Forensic_Software': forensic_software_text,
            'Report_Software': report_software_text,
            'Report_Article': report_article,
            'Processing_Software_List': processing_tools,
            'Report_Software_List': report_tools,
        }

        # Optional time frame if checked
        if self.time_frame_limited_var.get() == 1:
            data['PY_LIMITSTART'] = self.parse_request_date(self.time_frame_start_date.get().strip())
            data['PY_LIMITEND'] = self.parse_request_date(self.time_frame_end_date.get().strip())
            data['Time_Frame_Start'] = data['PY_LIMITSTART']
            data['Time_Frame_End'] = data['PY_LIMITEND']
        else:
            data['PY_LIMITSTART'] = ""
            data['PY_LIMITEND'] = ""
            data['Time_Frame_Start'] = ""
            data['Time_Frame_End'] = ""

        # Create search docs for replacement
        search_docs = {}
        for key in ['PY_DFR', 'PY_CASENUMBER', 'PY_EXAMINER', 'PY_REQAGENCY', 'PY_REQOFF', 'PY_SERVEDATE', 'PY_RETURNDATE', 'PY_PROVIDER', 'PY_ACCOUNTID', 'PY_DATASIZE', 'PY_OWNER', 'PY_LIMITSTART', 'PY_LIMITEND']:
            search_docs[key] = Document()
            p = search_docs[key].add_paragraph(data.get(key, ''))
            p.style = search_docs[key].styles['Normal']

        summary_doc = Document()
        self.generate_paragraphs(data, summary_doc)
        search_docs['PY_TEXT'] = summary_doc

        doc = Document(self.template_file)
        self.search_and_replace_content_controls_simple(doc, search_docs)
        self.search_and_replace_split_placeholders(doc, search_docs)

        # Save
        self.save_document(doc)

    def generate_replacements(self, data, search_docs):
        def add_text(doc, text):
            p = doc.add_paragraph(text)
            for run in p.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(11)
        for key, doc in search_docs.items():
            if key == "PY_TEXT":
                continue
            add_text(doc, data.get(key, ''))

    def save_document(self, doc):
        apply_warrant_suggested_filename(self)
        filename = ""
        try:
            filename = self.output_filename.get().strip()
        except Exception:
            filename = ""
        if not filename:
            filename = apply_warrant_suggested_filename(self)
        if filename.lower().endswith(".docx") is False:
            filename += ".docx"

        # Get default save directory from the UI Entry widget (or fallback)
        default_dir = None
        if hasattr(self, 'save_location'):
            default_dir = self.save_location.get().strip()
        if not default_dir:
            default_dir = self.current_settings.get("default_save_path", os.path.expanduser("~/Desktop"))

        # Full default path
        default_path = os.path.join(default_dir, filename)

        # Auto-save if valid directory exists
        if default_dir and os.path.isdir(default_dir):
            try:
                final_path = default_path
                # Avoid overwrite by adding suffix if file exists
                if os.path.exists(final_path):
                    base, ext = os.path.splitext(final_path)
                    i = 1
                    while os.path.exists(final_path):
                        final_path = f"{base}_{i}{ext}"
                        i += 1

                doc.save(final_path)
                remember_titles_from_form(self)
                remember_agencies_from_form(self)
                messagebox.showinfo("Success", f"Report auto-saved to:\n{final_path}")

                # Optional: remember this directory for next time
                self.current_settings["default_save_path"] = default_dir
                self.settings_manager.save_settings(self.current_settings)
                return
            except Exception as e:
                messagebox.showwarning("Auto-save Failed", f"Auto-save error:\n{str(e)}\n\nFalling back to manual selection.")
                # Continue to dialog

        # Fallback: manual save dialog
        save_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            initialdir=default_dir,
            initialfile=filename,
            title="Save Warrant Data Return Report"
        )

        if save_path:
            try:
                doc.save(save_path)
                remember_titles_from_form(self)
                remember_agencies_from_form(self)
                messagebox.showinfo("Success", f"Report saved to:\n{save_path}")

                # Remember the chosen directory
                chosen_dir = os.path.dirname(save_path)
                self.current_settings["default_save_path"] = chosen_dir
                self.settings_manager.save_settings(self.current_settings)
            except Exception as e:
                messagebox.showerror("Save Error", f"Could not save:\n{str(e)}")
        else:
            messagebox.showinfo("Cancelled", "Save cancelled.")

    def format_agency(self, agency, return_abbreviation=False):
        if not agency:
            return "" if not return_abbreviation else ("", "")
        
        agencies = {
            "federal bureau of investigation": ("Federal Bureau of Investigation", "FBI"),
            "fbi": ("Federal Bureau of Investigation", "FBI"),
            "drug enforcement administration": ("Drug Enforcement Administration", "DEA"),
            "dea": ("Drug Enforcement Administration", "DEA"),
            "bureau of alcohol tobacco firearms and explosives": ("Bureau of Alcohol Tobacco Firearms and Explosives", "ATF"),
            "atf": ("Bureau of Alcohol Tobacco Firearms and Explosives", "ATF"),
            "department of homeland security": ("Department of Homeland Security", "DHS"),
            "dhs": ("Department of Homeland Security", "DHS"),
            "united states marshals service": ("United States Marshals Service", "USMS"),
            "usms": ("United States Marshals Service", "USMS"),
            "south dakota division of criminal investigation": ("South Dakota Division of Criminal Investigation", "DCI"),
            "dci": ("South Dakota Division of Criminal Investigation", "DCI"),
            "south dakota dci": ("South Dakota Division of Criminal Investigation", "DCI"),
            "bureau of indian affairs": ("Bureau of Indian Affairs", "BIA"),
            "bia": ("Bureau of Indian Affairs", "BIA"),
            "south dakota highway patrol": ("South Dakota Highway Patrol", "SDHP"),
            "sdhp": ("South Dakota Highway Patrol", "SDHP"),
            "national security agency": ("National Security Agency", "NSA"),
            "nsa": ("National Security Agency", "NSA"),
            "federal emergency management agency": ("Federal Emergency Management Agency", "FEMA"),
            "fema": ("Federal Emergency Management Agency", "FEMA"),
            "internal revenue service": ("Internal Revenue Service", "IRS"),
            "irs": ("Internal Revenue Service", "IRS"),
        }
        
        agency_lower = agency.lower().strip()
        
        if agency_lower in agencies:
            full_name, abbr = agencies[agency_lower]
            return (full_name, abbr) if return_abbreviation else full_name
        
        original_agency = agency.strip()
        original_lower = original_agency.lower()
        
        abbreviation = None
        if original_lower.endswith(" pd"):
            location = original_agency[:-3].strip()
            agency = f"{location} Police Department"
            words = location.split()
            abbr_chars = [word[0].upper() for word in words if word]
            abbreviation = ''.join(abbr_chars) + "PD"
        elif original_lower.endswith(" so"):
            location = original_agency[:-3].strip()
            agency = f"{location} Sheriff's Office"
            words = location.split()
            abbr_chars = [word[0].upper() for word in words if word]
            abbreviation = ''.join(abbr_chars) + "SO"
        
        formatted_name = title_agency_words(agency)
        if len(agency) <= 4:
            formatted_name = agency.upper()
        
        if abbreviation is None:
            found_match = False
            for full_lower, (full_name, abbr) in agencies.items():
                if full_lower in agency_lower:
                    formatted_name = full_name
                    abbreviation = abbr
                    found_match = True
                    break
            if not found_match:
                if len(agency) <= 4:
                    abbreviation = agency.upper()
                else:
                    words = agency.split()
                    skip_words = ['of', 'the', 'and', 'in', 'on', 'at', 'by', 'for', 'with', 'a', 'an']
                    abbr_chars = [word[0].upper() for word in words if word.lower() not in skip_words and word]
                    abbreviation = ''.join(abbr_chars)
        
        if return_abbreviation:
            return formatted_name, abbreviation
        return formatted_name

    def format_title(self, title):
        if len(title) <= 2:
            return title.upper()
        return ' '.join(word.capitalize() for word in title.split())

    def get_request_agency_formatted(self):
        agency_text = self.requesting_agency.get().strip()
        if agency_text:
            return self.format_agency(agency_text)
        return ""

    def get_request_title(self):
        return (self.requesting_officer_title.get() or "").strip()

    def get_examiner_agency(self):
        agency = (self.examiner_agency.get() or "").strip()
        if agency == "South Dakota DCI":
            return "South Dakota Division of Criminal Investigation"
        return agency

    def parse_request_date(self, date_string):
        return parse_mdy_date(date_string, empty_ok=True)

    def format_time_frame_date(self, date_string):
        return self.parse_request_date(date_string)

    def search_and_replace_content_controls_simple(self, doc, search_docs):
        replaced_strings = {}
        for search_string in search_docs.keys():
            replaced_strings[search_string] = False

        doc_xml_str = etree.tostring(doc._element, encoding='unicode')

        for search_string, replacement_doc in search_docs.items():
            if search_string in doc_xml_str:
                if search_string == "PY_TEXT":
                    replacement_xml_parts = []
                    for para in replacement_doc.paragraphs:
                        para_text = para.text or ""
                        text_parts = para_text.split('\n')
                        for text_part in text_parts:
                            if not text_part and len(text_parts) > 1:
                                replacement_xml_parts.append('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr></w:pPr></w:p>')
                                continue
                            para_xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr></w:pPr>'
                            if text_part:
                                para_xml += f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">{text_part.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")}</w:t></w:r>'
                            para_xml += '</w:p>'
                            replacement_xml_parts.append(para_xml)
                    replacement_text = ''.join(replacement_xml_parts)
                else:
                    replacement_text = replacement_doc.paragraphs[0].text if replacement_doc.paragraphs else ""
                    replacement_text = replacement_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

                doc_xml_str = doc_xml_str.replace(search_string, replacement_text)
                replaced_strings[search_string] = True

        new_element = etree.fromstring(doc_xml_str.encode('utf-8'))
        doc._element.clear()
        for child in new_element:
            doc._element.append(child)
        
        return replaced_strings

    def search_and_replace_split_placeholders(self, doc, search_docs):
        replaced_strings = {}
        for search_string in search_docs.keys():
            replaced_strings[search_string] = False

        try:
            doc_xml_str = etree.tostring(doc._element, encoding='unicode')
            original_xml_str = doc_xml_str

            target_placeholders = ['PY_DFR', 'PY_EXAMINER', 'PY_REQAGENCY', 'PY_REQOFF', 'PY_SERVEDATE', 'PY_RETURNDATE', 'PY_PROVIDER', 'PY_ACCOUNTID', 'PY_DATASIZE', 'PY_OWNER']

            for search_string in target_placeholders:
                if replaced_strings.get(search_string):
                    continue
                replacement_doc = search_docs.get(search_string)
                if not replacement_doc or not replacement_doc.paragraphs:
                    continue
                replacement_text = replacement_doc.paragraphs[0].text or ""
                replacement_text = replacement_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&apos;')

                pattern_parts = [re.escape(char) for char in search_string]
                pattern = r'(?:<[^>]*>)*'.join(pattern_parts)

                matches = list(re.finditer(pattern, doc_xml_str))
                if matches:
                    for match in reversed(matches):
                        doc_xml_str = doc_xml_str[:match.start()] + replacement_text + doc_xml_str[match.end():]
                    replaced_strings[search_string] = True

            if doc_xml_str != original_xml_str:
                new_element = etree.fromstring(doc_xml_str.encode('utf-8'))
                doc._element.clear()
                for child in new_element:
                    doc._element.append(child)
        except:
            pass
        return replaced_strings

    def on_closing(self):
        close_and_return(self)

if __name__ == "__main__":
    app = WarrantDataReturns()
    app.mainloop()
    
# Ω Digital Forensics Report Writer Ω (ver. 1.0.3) © 2026 #