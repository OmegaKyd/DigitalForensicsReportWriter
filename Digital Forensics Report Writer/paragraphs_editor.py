"""Tools → Edit Paragraphs window."""
import re
import tkinter as tk
from tkinter import messagebox, ttk

from paragraphs_manager import (
    REPORT_TYPES,
    TOKEN_HELP,
    default_paragraphs,
    fill_paragraph,
    is_modified,
    load_paragraphs,
    paragraph_label,
    paragraph_when,
    paragraphs_file,
    refresh_open_report_windows,
    revert_all,
    revert_paragraph,
    revert_report,
    save_paragraph,
    tokens_for_report,
    validate_paragraph_text,
)
from ui_theme import COLORS

_TOKEN_SPAN = re.compile(r"\{[A-Za-z_][A-Za-z0-9_]*\}")


def open_paragraph_editor(parent):
    win = tk.Toplevel(parent)
    win.title("Edit Paragraphs")
    win.configure(bg=COLORS["bg"])
    win.geometry("1100x680")
    win.minsize(900, 560)
    win.transient(parent)

    current_kind = tk.StringVar(value=REPORT_TYPES[0][0])
    current_key = {"value": None}
    loaded_text = {"value": ""}
    loading = {"value": False}
    refreshing = {"value": False}

    ttk.Label(
        win,
        text="Edit the canned narrative for each report type. Field tokens such as {Request_Date} "
        "are filled from the form when a report is generated. PY_ tokens belong in Word templates, "
        "not here (Help → Template Placeholders).",
        wraplength=1060,
        style="Hint.TLabel",
    ).pack(anchor="w", padx=12, pady=(10, 4))

    body = ttk.Frame(win)
    body.pack(fill="both", expand=True, padx=12, pady=(0, 8))
    body.columnconfigure(2, weight=1)
    body.rowconfigure(0, weight=1)

    left = ttk.Frame(body)
    left.grid(row=0, column=0, sticky="nsw", padx=(0, 8))
    ttk.Label(left, text="Report type").pack(anchor="w")
    type_list = tk.Listbox(
        left,
        height=8,
        width=28,
        exportselection=False,
        bg=COLORS["entry_bg"],
        fg=COLORS["text"],
        selectbackground=COLORS["accent_dark"],
        highlightthickness=0,
    )
    type_list.pack(fill="y", pady=(4, 10))
    for _kind, label in REPORT_TYPES:
        type_list.insert("end", label)
    type_list.selection_set(0)

    ttk.Label(left, text="Paragraph").pack(anchor="w")
    para_frame = ttk.Frame(left)
    para_frame.pack(fill="both", expand=True)
    para_scroll = ttk.Scrollbar(para_frame)
    para_scroll.pack(side="right", fill="y")
    para_list = tk.Listbox(
        para_frame,
        width=36,
        exportselection=False,
        yscrollcommand=para_scroll.set,
        bg=COLORS["entry_bg"],
        fg=COLORS["text"],
        selectbackground=COLORS["accent_dark"],
        highlightthickness=0,
    )
    para_list.pack(side="left", fill="both", expand=True)
    para_scroll.config(command=para_list.yview)

    mid = ttk.Frame(body)
    mid.grid(row=0, column=2, sticky="nsew")
    mid.rowconfigure(3, weight=1)
    mid.columnconfigure(0, weight=1)

    title_label = ttk.Label(mid, text="", font=("Segoe UI", 11, "bold"))
    title_label.grid(row=0, column=0, sticky="w")
    when_label = ttk.Label(mid, text="", style="Hint.TLabel", wraplength=620)
    when_label.grid(row=1, column=0, sticky="w", pady=(2, 2))
    status_label = ttk.Label(mid, text="", style="Hint.TLabel")
    status_label.grid(row=2, column=0, sticky="w", pady=(0, 4))

    editor = tk.Text(
        mid,
        wrap="word",
        undo=True,
        bg=COLORS["entry_bg"],
        fg=COLORS["text"],
        insertbackground=COLORS["accent"],
        relief="flat",
        font=("Segoe UI", 10),
    )
    editor.grid(row=3, column=0, sticky="nsew")
    editor.tag_configure(
        "token_ok",
        foreground=COLORS["accent"],
        font=("Segoe UI", 10, "bold"),
    )
    editor.tag_configure(
        "token_bad",
        foreground=COLORS["warn"],
        font=("Segoe UI", 10, "bold"),
    )
    highlight_job = {"id": None}

    preview_label = ttk.Label(
        mid,
        text="Known field tokens turn cyan as soon as you finish typing {Name}. "
        "Amber means that name is not used by this report type.",
        style="Hint.TLabel",
        wraplength=620,
    )
    preview_label.grid(row=4, column=0, sticky="w", pady=(6, 0))

    right = ttk.Frame(body, width=260)
    right.grid(row=0, column=3, sticky="nsw", padx=(10, 0))
    ttk.Label(right, text="Field tokens").pack(anchor="w")
    ttk.Label(
        right,
        text="Double-click to insert at the cursor.",
        style="Hint.TLabel",
        wraplength=240,
    ).pack(anchor="w", pady=(0, 4))
    token_box = ttk.Frame(right)
    token_box.pack(fill="both", expand=True)
    token_scroll = ttk.Scrollbar(token_box)
    token_list = tk.Listbox(
        token_box,
        width=34,
        height=22,
        exportselection=False,
        yscrollcommand=token_scroll.set,
        bg=COLORS["entry_bg"],
        fg=COLORS["text"],
        selectbackground=COLORS["accent_dark"],
        highlightthickness=0,
        font=("Consolas", 9),
    )
    token_list.pack(side="left", fill="both", expand=True)
    token_scroll.pack(side="right", fill="y")
    token_scroll.config(command=token_list.yview)
    token_help = ttk.Label(right, text="", style="Hint.TLabel", wraplength=240)
    token_help.pack(anchor="w", pady=(8, 0))

    keys_for_kind = {"items": []}

    def selected_kind():
        selection = type_list.curselection()
        if not selection:
            return REPORT_TYPES[0][0]
        return REPORT_TYPES[selection[0]][0]

    def selected_key():
        selection = para_list.curselection()
        if not selection:
            return None
        index = selection[0]
        if 0 <= index < len(keys_for_kind["items"]):
            return keys_for_kind["items"][index]
        return None

    def current_editor_text():
        return editor.get("1.0", "end-1c")

    def paragraph_is_dirty():
        if loading["value"] or refreshing["value"]:
            return False
        return current_editor_text() != loaded_text["value"]

    def confirm_leave():
        if not paragraph_is_dirty():
            return True
        answer = messagebox.askyesnocancel(
            "Unsaved changes",
            "Save changes to this paragraph before leaving?",
            parent=win,
        )
        if answer is None:
            return False
        if answer:
            return save_current(quiet=True)
        editor.delete("1.0", "end")
        editor.insert("1.0", loaded_text["value"])
        editor.edit_modified(False)
        highlight_tokens()
        return True

    def fill_paragraph_list(kind, prefer_key=None):
        refreshing["value"] = True
        keys_for_kind["items"] = list(default_paragraphs(kind).keys())
        para_list.delete(0, "end")
        pick = 0
        for index, key in enumerate(keys_for_kind["items"]):
            flag = " ●" if is_modified(kind, key) else ""
            para_list.insert("end", f"{paragraph_label(kind, key)}{flag}")
            if prefer_key and key == prefer_key:
                pick = index
        if keys_for_kind["items"]:
            para_list.selection_clear(0, "end")
            para_list.selection_set(pick)
            para_list.see(pick)
        refreshing["value"] = False

    def known_token_strings(kind=None):
        kind = kind or current_kind.get()
        return {"{" + token + "}" for token in tokens_for_report(kind)}

    def highlight_tokens():
        highlight_job["id"] = None
        known = known_token_strings()
        editor.tag_remove("token_ok", "1.0", "end")
        editor.tag_remove("token_bad", "1.0", "end")
        text = editor.get("1.0", "end-1c")
        for match in _TOKEN_SPAN.finditer(text):
            start = f"1.0+{match.start()}c"
            end = f"1.0+{match.end()}c"
            tag = "token_ok" if match.group() in known else "token_bad"
            editor.tag_add(tag, start, end)

    def schedule_highlight(_event=None):
        job = highlight_job.get("id")
        if job is not None:
            try:
                win.after_cancel(job)
            except Exception:
                pass
        highlight_job["id"] = win.after(30, highlight_tokens)

    def fill_tokens(kind):
        token_list.delete(0, "end")
        for token in tokens_for_report(kind):
            token_list.insert("end", "{" + token + "}")
            token_list.itemconfig(token_list.size() - 1, fg=COLORS["accent"])

    def show_current():
        kind = selected_kind()
        key = selected_key()
        current_kind.set(kind)
        current_key["value"] = key
        loading["value"] = True
        editor.delete("1.0", "end")
        if not key:
            title_label.configure(text="")
            when_label.configure(text="")
            status_label.configure(text="")
            loaded_text["value"] = ""
            loading["value"] = False
            editor.edit_modified(False)
            return
        text = load_paragraphs(kind).get(key, "")
        editor.insert("1.0", text)
        loaded_text["value"] = text
        title_label.configure(text=paragraph_label(kind, key))
        when_label.configure(text=paragraph_when(kind, key) or "")
        if is_modified(kind, key):
            status_label.configure(text="Modified from the factory default.")
        else:
            status_label.configure(text="Factory default.")
        loading["value"] = False
        editor.edit_modified(False)
        highlight_tokens()

    def save_current(quiet=False):
        kind = current_kind.get()
        key = current_key["value"]
        if not key:
            return True
        if not paragraph_is_dirty():
            return True
        text = current_editor_text()
        problems = validate_paragraph_text(text)
        if problems:
            messagebox.showerror("Invalid tokens", "\n".join(problems), parent=win)
            return False
        if not save_paragraph(kind, key, text):
            messagebox.showerror("Save failed", "Could not write paragraphs.json.", parent=win)
            return False
        loaded_text["value"] = text
        fill_paragraph_list(kind, prefer_key=key)
        show_current()
        refresh_open_report_windows(parent)
        if not quiet:
            status_label.configure(text=f"Saved. Stored in {paragraphs_file()}")
        return True

    def on_type_change(_event=None):
        if refreshing["value"] or loading["value"]:
            return
        if selected_kind() == current_kind.get():
            return
        if not confirm_leave():
            # restore previous type selection
            wanted = current_kind.get()
            for index, (kind, _label) in enumerate(REPORT_TYPES):
                if kind == wanted:
                    type_list.selection_clear(0, "end")
                    type_list.selection_set(index)
                    break
            return
        kind = selected_kind()
        fill_tokens(kind)
        fill_paragraph_list(kind)
        show_current()

    def on_para_change(_event=None):
        if refreshing["value"] or loading["value"]:
            return
        key = selected_key()
        if key == current_key["value"]:
            return
        if not confirm_leave():
            kind = current_kind.get()
            fill_paragraph_list(kind, prefer_key=current_key["value"])
            return
        show_current()

    def insert_token(_event=None):
        selection = token_list.curselection()
        if not selection:
            return
        token = token_list.get(selection[0])
        editor.insert("insert", token)
        editor.focus_set()
        highlight_tokens()

    def show_token_help(_event=None):
        selection = token_list.curselection()
        if not selection:
            token_help.configure(text="")
            return
        raw = token_list.get(selection[0]).strip("{}")
        token_help.configure(text=TOKEN_HELP.get(raw, "Filled from the report form or extraction file."))

    def preview_sample():
        kind = current_kind.get()
        text = current_editor_text()
        sample = {token: token.replace("_", " ").title() for token in tokens_for_report(kind)}
        filled = fill_paragraph(text, sample)
        preview_label.configure(text="Sample fill: " + " ".join(filled.split())[:240])

    def do_revert_this():
        kind = current_kind.get()
        key = current_key["value"]
        if not key:
            return
        if not messagebox.askyesno(
            "Revert paragraph",
            "Replace this paragraph with the factory default?",
            parent=win,
        ):
            return
        revert_paragraph(kind, key)
        fill_paragraph_list(kind, prefer_key=key)
        show_current()
        refresh_open_report_windows(parent)

    def do_revert_type():
        kind = current_kind.get()
        if not messagebox.askyesno(
            "Revert report type",
            f"Restore every paragraph in {type_list.get(type_list.curselection()[0])} to the factory default?",
            parent=win,
        ):
            return
        revert_report(kind)
        fill_paragraph_list(kind, prefer_key=current_key["value"])
        show_current()
        refresh_open_report_windows(parent)

    def do_revert_all():
        if not messagebox.askyesno(
            "Revert all paragraphs",
            "Restore every report type to the factory defaults and delete saved paragraph edits?",
            parent=win,
        ):
            return
        revert_all()
        fill_paragraph_list(current_kind.get(), prefer_key=current_key["value"])
        show_current()
        refresh_open_report_windows(parent)

    def on_close():
        if not confirm_leave():
            return
        win.destroy()

    editor.bind("<<Modified>>", lambda e: (editor.edit_modified(False), schedule_highlight()))
    editor.bind("<KeyRelease>", schedule_highlight)
    editor.bind("<<Paste>>", schedule_highlight)
    type_list.bind("<<ListboxSelect>>", on_type_change)
    para_list.bind("<<ListboxSelect>>", on_para_change)
    token_list.bind("<<ListboxSelect>>", show_token_help)
    token_list.bind("<Double-Button-1>", insert_token)

    buttons = ttk.Frame(win)
    buttons.pack(fill="x", padx=12, pady=(0, 12))
    ttk.Button(buttons, text="Save", command=save_current).pack(side="left", padx=4)
    ttk.Button(buttons, text="Preview Sample", command=preview_sample).pack(side="left", padx=4)
    ttk.Button(buttons, text="Revert This Paragraph", command=do_revert_this).pack(side="left", padx=4)
    ttk.Button(buttons, text="Revert This Report Type", command=do_revert_type).pack(side="left", padx=4)
    ttk.Button(buttons, text="Revert All to Defaults", command=do_revert_all).pack(side="left", padx=4)
    ttk.Button(buttons, text="Close", command=on_close).pack(side="right", padx=4)

    fill_tokens(selected_kind())
    fill_paragraph_list(selected_kind())
    show_current()
    win.protocol("WM_DELETE_WINDOW", on_close)
    win.grab_set()
    win.focus_set()
