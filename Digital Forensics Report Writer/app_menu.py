"""Shared File / Tools / Help menu for report windows and the start screen."""
import os
import shutil
import sys
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from ui_theme import (
    APP_NAME,
    APP_VERSION,
    COPYRIGHT_YEAR,
    COLORS,
    GITHUB_URL,
    resource_dir,
    writable_dir,
)


PY_PLACEHOLDERS = [
    ("Case / request", [
        ("PY_DFR", "DFR number"),
        ("PY_CASENUMBER", "Agency or lab case number"),
        ("PY_REQDATE", "Request date (formatted)"),
        ("PY_REQOFF", "Requesting officer name"),
        ("PY_REQAGENCY", "Requesting agency"),
        ("PY_OWNER", "Device owner (mobile/PC) or account owner (warrant)"),
        ("PY_EVIDENCE", "Evidence number"),
    ]),
    ("Device", [
        ("PY_MAN / PY_DEVMAKE / PY_PCMAN", "Device or computer manufacturer"),
        ("PY_MOD / PY_DEVMODEL / PY_PCMOD", "Device or computer model"),
        ("PY_SERIAL / PY_PCSERIAL", "Device or computer serial number"),
        ("PY_DEVNAME", "Device name"),
        ("PY_COLOR", "Device color"),
        ("PY_OS", "Device operating system"),
        ("PY_IMEI", "IMEI (first listed)"),
        ("PY_ICCID", "ICCID (first listed)"),
        ("PY_PHONE", "Phone number"),
        ("PY_CARRIER", "Carrier"),
        ("PY_PASSCODE", "Passcode or lock status"),
        ("PY_CAPACITY", "Storage capacity"),
        ("PY_HDMAKE", "Hard-drive manufacturer"),
        ("PY_HDMODEL", "Hard-drive model"),
        ("PY_HDSERIAL", "Hard-drive serial"),
    ]),
    ("Acquisition / software", [
        ("PY_IMAGEDATE", "Acquisition or image date"),
        ("PY_ACQUIRE", "Acquisition tool / method text"),
        ("PY_CBVER", "Cellebrite version"),
        ("PY_GKVER", "GrayKey version"),
        ("PY_DCVER", "Cellebrite Digital Collector version (PC / Mac acquisitions)"),
        ("PY_FTKVER", "FTK Imager version"),
        ("PY_TX1VER", "TX1 OS / version"),
        ("PY_XWVER", "X-Ways version"),
    ]),
    ("Warrant / provider", [
        ("PY_PROVIDER", "Service provider"),
        ("PY_ACCOUNT / PY_ACCOUNTID", "Account identifier"),
        ("PY_DATASIZE", "Returned data size"),
        ("PY_OWNER", "Account owner"),
        ("PY_SERVEDATE", "Warrant service date"),
        ("PY_RETURNDATE", "Data return date"),
        ("PY_LIMITSTART", "Time-frame start"),
        ("PY_LIMITEND", "Time-frame end"),
        ("PY_EXAMINER", "Examiner name"),
    ]),
    ("Body slots", [
        ("PY_TEXT", "Inserted narrative paragraphs"),
    ]),
]


def user_templates_dir():
    from report_common import templates_dir
    return templates_dir()


def open_paragraph_editor(parent):
    from paragraphs_editor import open_paragraph_editor as _open
    _open(parent)


def open_titles_editor(parent):
    from titles_editor import open_titles_editor as _open
    _open(parent)


def open_agencies_editor(parent):
    from agencies_editor import open_agencies_editor as _open
    _open(parent)


def _change_template_folder(parent):
    from report_common import change_templates_location

    folder = change_templates_location(parent)
    if folder is None:
        return
    _refresh_open_pickers(parent)


def attach_app_menu(window, is_start=False):
    menubar = tk.Menu(window)
    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="Exit", command=lambda: _exit_program(window, is_start))
    menubar.add_cascade(label="File", menu=file_menu)

    tools_menu = tk.Menu(menubar, tearoff=0)
    tools_menu.add_command(label="Manage Templates...", command=lambda: open_template_manager(window))
    tools_menu.add_command(
        label="Change Template Folder Location...",
        command=lambda: _change_template_folder(window),
    )
    tools_menu.add_separator()
    tools_menu.add_command(label="Edit Paragraphs...", command=lambda: open_paragraph_editor(window))
    tools_menu.add_command(label="Edit Officer Titles...", command=lambda: open_titles_editor(window))
    tools_menu.add_command(label="Edit Agencies...", command=lambda: open_agencies_editor(window))
    menubar.add_cascade(label="Tools", menu=tools_menu)

    help_menu = tk.Menu(menubar, tearoff=0)
    help_menu.add_command(label="Template Placeholders (PY_)...", command=lambda: show_placeholder_guide(window))
    help_menu.add_separator()
    help_menu.add_command(label="About", command=lambda: show_about(window))
    menubar.add_cascade(label="Help", menu=help_menu)

    window.config(menu=menubar)
    return menubar


def _autosave_window(target):
    if target is None:
        return
    try:
        timer = getattr(target, "_save_timer", None)
        if timer:
            target.after_cancel(timer)
            target._save_timer = None
    except Exception:
        pass
    try:
        if hasattr(target, "_perform_auto_save"):
            target._perform_auto_save()
    except Exception:
        pass


def _exit_program(window, is_start):
    master = getattr(window, "master", None)
    _autosave_window(window)
    _autosave_window(master)
    try:
        window.destroy()
    except Exception:
        pass
    if master is not None:
        try:
            master.destroy()
        except Exception:
            pass
    try:
        sys.exit(0)
    except SystemExit:
        raise
    except Exception:
        os._exit(0)


def show_about(parent):
    win = tk.Toplevel(parent)
    win.title("About")
    win.configure(bg=COLORS["bg"])
    win.resizable(False, False)
    win.transient(parent)

    frame = ttk.Frame(win, padding=18)
    frame.pack(fill="both", expand=True)

    ttk.Label(frame, text=APP_NAME, font=("Segoe UI", 12, "bold")).pack(anchor="w")
    ttk.Label(frame, text=f"Version {APP_VERSION}").pack(anchor="w", pady=(6, 0))
    ttk.Label(frame, text=f"© {COPYRIGHT_YEAR}").pack(anchor="w", pady=(2, 10))

    link = tk.Label(
        frame,
        text=GITHUB_URL,
        fg=COLORS["accent"],
        bg=COLORS["bg"],
        font=("Segoe UI", 10, "underline"),
        cursor="hand2",
    )
    link.pack(anchor="w")
    link.bind("<Button-1>", lambda _event: webbrowser.open(GITHUB_URL))
    link.bind("<Enter>", lambda _event: link.configure(fg=COLORS["text"]))
    link.bind("<Leave>", lambda _event: link.configure(fg=COLORS["accent"]))

    ttk.Button(frame, text="Close", command=win.destroy).pack(anchor="e", pady=(16, 0))

    win.update_idletasks()
    width = max(360, win.winfo_reqwidth())
    height = win.winfo_reqheight()
    px = parent.winfo_rootx() + max(0, (parent.winfo_width() - width) // 2)
    py = parent.winfo_rooty() + max(0, (parent.winfo_height() - height) // 3)
    win.geometry(f"{width}x{height}+{px}+{py}")
    win.grab_set()
    win.focus_set()


def show_placeholder_guide(parent):
    win = tk.Toplevel(parent)
    win.title("Template Placeholders")
    win.configure(bg=COLORS["bg"])
    win.geometry("720x560")
    ttk.Label(
        win,
        text="Include these PY_ tokens in new Word templates. They are replaced when a report is generated.",
        wraplength=680,
    ).pack(anchor="w", padx=12, pady=(12, 6))
    text = tk.Text(
        win,
        wrap="word",
        bg=COLORS["entry_bg"],
        fg=COLORS["text"],
        insertbackground=COLORS["accent"],
        relief="flat",
    )
    text.pack(fill="both", expand=True, padx=12, pady=(0, 8))
    for title, rows in PY_PLACEHOLDERS:
        text.insert("end", title + "\n", "section")
        for token, meaning in rows:
            text.insert("end", f"  {token}\n", "token")
            text.insert("end", f"      {meaning}\n")
        text.insert("end", "\n")
    text.tag_configure("section", foreground=COLORS["accent"], font=("Segoe UI", 11, "bold"))
    text.tag_configure("token", font=("Consolas", 10, "bold"))
    text.config(state="disabled")
    ttk.Button(win, text="Close", command=win.destroy).pack(pady=(0, 10))


def _refresh_open_pickers(parent):
    from report_common import list_dfr_templates, default_template_for

    folder = user_templates_dir()
    for widget in (parent, getattr(parent, "master", None)):
        if widget is None:
            continue
        combo = getattr(widget, "template_choice", None)
        if combo is None:
            continue
        keywords = getattr(widget, "_template_keywords", None)
        items = list_dfr_templates(keywords)
        choices = {item["name"]: item["path"] for item in items}
        widget._template_choices = choices
        try:
            combo.configure(values=list(choices.keys()))
            current = combo.get()
            if current not in choices:
                preferred = getattr(widget, "_template_preferred", "")
                picked = default_template_for(preferred, keywords)
                if picked:
                    combo.set(picked["name"])
                    widget.template_file = picked["path"]
                elif choices:
                    name = next(iter(choices))
                    combo.set(name)
                    widget.template_file = choices[name]
                else:
                    combo.set("")
                    widget.template_file = None
        except Exception:
            pass
    return folder


def open_template_manager(parent):
    folder = user_templates_dir()
    win = tk.Toplevel(parent)
    win.title("Manage Templates")
    win.configure(bg=COLORS["bg"])
    win.geometry("720x460")
    win.minsize(560, 400)
    win.transient(parent)

    path_label = ttk.Label(win, text=f"Folder: {folder}", wraplength=680)
    path_label.pack(anchor="w", padx=12, pady=(10, 4))
    ttk.Label(
        win,
        text="Add, rename, or remove .docx templates. Files are stored in your DFR Templates folder. "
             "Removed files stay removed after restart. Use Restore Official Templates if you want the packaged files back.",
        wraplength=680,
        style="Hint.TLabel",
    ).pack(anchor="w", padx=12, pady=(0, 8))

    listbox = tk.Listbox(
        win,
        bg=COLORS["entry_bg"],
        fg=COLORS["text"],
        selectbackground=COLORS["accent_dark"],
        highlightthickness=0,
    )
    listbox.pack(fill="both", expand=True, padx=12, pady=(0, 8))

    def current_folder():
        from report_common import resolve_templates_dir
        return resolve_templates_dir()

    def refresh():
        folder = current_folder()
        path_label.configure(text=f"Folder: {folder}")
        listbox.delete(0, "end")
        try:
            files = sorted(folder.glob("*.docx"))
        except Exception:
            files = []
        for path in files:
            if not path.name.startswith("~$"):
                listbox.insert("end", path.name)

    def selected_path():
        selection = listbox.curselection()
        if not selection:
            return None
        return current_folder() / listbox.get(selection[0])

    def add_template():
        source = filedialog.askopenfilename(
            parent=win,
            title="Add template",
            filetypes=[("Word documents", "*.docx")],
        )
        if not source:
            return
        dest = current_folder() / Path(source).name
        if dest.exists():
            if not messagebox.askyesno("Replace?", f"{dest.name} already exists. Replace it?", parent=win):
                return
        shutil.copy2(source, dest)
        refresh()
        _refresh_open_pickers(parent)

    def rename_template():
        path = selected_path()
        if path is None:
            messagebox.showinfo("Manage Templates", "Select a template first.", parent=win)
            return
        new_name = simpledialog.askstring("Rename Template", "New file name:", initialvalue=path.stem, parent=win)
        if not new_name:
            return
        new_name = new_name.strip()
        if not new_name.lower().endswith(".docx"):
            new_name += ".docx"
        dest = current_folder() / new_name
        if dest.exists():
            messagebox.showerror("Rename", "A template with that name already exists.", parent=win)
            return
        path.rename(dest)
        refresh()
        _refresh_open_pickers(parent)

    def remove_template():
        path = selected_path()
        if path is None:
            messagebox.showinfo("Manage Templates", "Select a template first.", parent=win)
            return
        if not messagebox.askyesno("Remove Template", f"Delete {path.name} from the templates folder?", parent=win):
            return
        try:
            if path.exists():
                path.unlink()
            else:
                messagebox.showerror(
                    "Remove Template",
                    f"{path.name} was not found in:\n{path.parent}",
                    parent=win,
                )
                refresh()
                return
        except Exception as exc:
            messagebox.showerror(
                "Remove Template",
                f"Could not delete {path.name}.\n\n{exc}\n\n"
                "Close the file if it is open in Word and try again.",
                parent=win,
            )
            return
        refresh()
        _refresh_open_pickers(parent)

    def restore_official():
        from report_common import restore_official_templates

        folder = current_folder()
        if not messagebox.askyesno(
            "Restore Official Templates",
            "Copy the packaged DFR templates back into this folder?\n\n"
            "Existing files with the same name will be left alone.",
            parent=win,
        ):
            return
        copied = restore_official_templates(folder, overwrite=False)
        refresh()
        _refresh_open_pickers(parent)
        if copied:
            messagebox.showinfo(
                "Restore Official Templates",
                "Restored:\n" + "\n".join(copied),
                parent=win,
            )
        else:
            messagebox.showinfo(
                "Restore Official Templates",
                "No packaged templates were missing from this folder.",
                parent=win,
            )

    def open_folder():
        try:
            os.startfile(str(current_folder()))
        except Exception:
            messagebox.showinfo("DFR Templates", str(current_folder()), parent=win)

    actions = ttk.Frame(win)
    actions.pack(fill="x", padx=12, pady=(0, 12))
    row1 = ttk.Frame(actions)
    row1.pack(fill="x")
    row2 = ttk.Frame(actions)
    row2.pack(fill="x", pady=(8, 0))
    ttk.Button(row1, text="Add...", command=add_template).pack(side="left", padx=(0, 6))
    ttk.Button(row1, text="Rename...", command=rename_template).pack(side="left", padx=(0, 6))
    ttk.Button(row1, text="Remove", command=remove_template).pack(side="left", padx=(0, 6))
    ttk.Button(row1, text="Open Folder", command=open_folder).pack(side="left")
    ttk.Button(row2, text="Restore Official Templates", command=restore_official).pack(side="left")
    ttk.Button(row2, text="Close", command=win.destroy).pack(side="right")
    refresh()
