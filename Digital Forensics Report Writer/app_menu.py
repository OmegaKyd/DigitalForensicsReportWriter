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
        ("PY_DFR", "DFR report number"),
        ("PY_CASENUMBER", "Agency or lab case number"),
        ("PY_EVIDENCE", "Evidence number"),
        ("PY_REQDATE", "Request date"),
        ("PY_OWNER", "Device owner, or warrant account owner"),
        ("PY_REQAGENCY", "Requesting / case agency"),
        ("PY_REQOFF", "Requesting officer or case agent"),
        ("PY_EXAMINER", "Examiner title and name"),
        ("PY_IMAGEDATE", "Exam start date and/or image date (same value if tagged twice)"),
    ]),
    ("Authority and time frame", [
        ("PY_AUTHCONSENT", "Checked when legal authority is Consent"),
        ("PY_AUTHSW", "Checked when legal authority is Search Warrant"),
        ("PY_AUTHIMPLIED", "Checked when legal authority is Implied Consent"),
        ("PY_AUTHPAROLE", "Checked when legal authority is Parole"),
        ("PY_AUTHOTHER", "Other-authority text slot (no program field yet)"),
        ("PY_TIMEFRAMEYES", "Checked when Time Frame Limited is Yes"),
        ("PY_TIMEFRAMENO", "Checked when Time Frame Limited is No"),
        ("PY_LIMITSTART", "Time-frame start date"),
        ("PY_LIMITEND", "Time-frame end date"),
    ]),
    ("Mobile / device identity", [
        ("PY_MAN", "Mobile manufacturer"),
        ("PY_MOD", "Mobile model"),
        ("PY_COLOR", "Color"),
        ("PY_PHONE", "Phone number"),
        ("PY_SERIAL", "Mobile serial number"),
        ("PY_IMEI", "IMEI (first listed)"),
        ("PY_CAPACITY", "Capacity"),
        ("PY_DEVNAME", "Device name"),
        ("PY_ACCOUNT", "Device account"),
        ("PY_ICCID", "ICCID (first listed)"),
        ("PY_PASSCODE", "Passcode text"),
        ("PY_CARRIER", "Carrier"),
        ("PY_OS", "Mobile OS version"),
        ("PY_AIRPLANEYES", "Checked when Airplane Mode When Received is Yes"),
        ("PY_AIRPLANENO", "Checked when Airplane Mode When Received is No"),
    ]),
    ("Computer / storage identity", [
        ("PY_DEVMAKE", "Computer or media manufacturer"),
        ("PY_DEVMODEL", "Computer or media model"),
        ("PY_PCSERIAL", "Computer serial number"),
        ("PY_OSVERSION", "Computer OS version"),
        ("PY_HDMAKE", "Hard-drive manufacturer"),
        ("PY_HDMODEL", "Hard-drive model"),
        ("PY_HDSERIAL", "Hard-drive serial"),
    ]),
    ("Acquisition versions and image checkboxes", [
        ("PY_CBVER", "Cellebrite version"),
        ("PY_CBYES", "Checked when Cellebrite is selected (mobile)"),
        ("PY_GKVER", "GrayKey version"),
        ("PY_FTKVER", "FTK Imager version"),
        ("PY_FTKYES", "Checked when the loaded log is FTK"),
        ("PY_TX1VER", "TX1 version"),
        ("PY_TX1YES", "Checked when the loaded log is TX1"),
        ("PY_XWVER", "X-Ways version"),
        ("PY_XWYES", "Checked when the loaded log is X-Ways (forensic image)"),
        ("PY_DCVER", "Digital Collector version"),
        ("PY_DCYES", "Checked when the loaded log is Digital Collector"),
        ("PY_ACQUIRE", "Image acquisition log / method text"),
    ]),
    ("Processing software", [
        ("PY_EXAMINEVER", "Magnet AXIOM version from the loaded HTML or XML export"),
        ("PY_EXAMINE", "Same version text on older templates"),
        ("PY_EXAMINEYES", "Checked when AXIOM / Magnet Examine is selected"),
        ("PY_XWPROYES", "Checked when X-Ways processing is selected"),
        ("PY_GRIFFEYEYES", "Checked when Griffeye is selected"),
        ("PY_CELLEBRITEYES", "Checked when Cellebrite is selected (warrant)"),
        ("PY_MANUAL", "Checked when Manual Exam Only is selected (warrant)"),
        ("PY_MANUALYES", "Same as PY_MANUAL if that tag is used"),
    ]),
    ("Warrant / provider", [
        ("PY_PROVIDER", "Service provider"),
        ("PY_ACCOUNTID", "Account identifier"),
        ("PY_DATASIZE", "Returned data size"),
        ("PY_SERVEDATE", "Warrant service date"),
        ("PY_RETURNDATE", "Data return date"),
    ]),
    ("Narrative", [
        ("PY_TEXT", "Canned narrative paragraphs"),
        ("PY_NOTES", "Notes tab text"),
    ]),
    ("Android prep checklist", [
        ("PY_ANDROID_DEV", "Enabled Developer Options"),
        ("PY_ANDROID_DEBUG", "Enabled USB Debugging"),
        ("PY_ANDROID_AWAKE", "Set to Stay Awake"),
        ("PY_ANDROID_TRANSFER", "USB Settings Set to File Transfer or MTP"),
        ("PY_ANDROID_VERIFY", "Verify Apps Turned Off"),
        ("PY_ANDROID_TIMEOUT", "Screen Timeout Set to Longest Setting"),
        ("PY_ANDROID_SOURCES", "Unknown Sources Selected in Developer Options"),
        ("PY_ANDROID_LOCKOFF", "Turned Lock Off"),
        ("PY_ANDROID_RECOVERY", "Enabled Recovery Mode"),
    ]),
    ("Apple iOS prep checklist", [
        ("PY_IOS_DEV", "Enabled Developer Options"),
        ("PY_IOS_SDP", "Turned Off Stolen Device Protection"),
        ("PY_IOS_TRUST", "Trust Computer"),
        ("PY_IOS_TIMEOUT", "Screen Timeout Set to Never"),
        ("PY_IOS_LOWPOWER", "Turned Off Low Power Mode"),
        ("PY_IOS_LOCKOFF", "Turned Off Lock"),
        ("PY_IOS_RECOVERY", "Entered Recovery Mode"),
        ("PY_IOS_DFU", "Entered DFU Mode"),
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
    if not is_start:
        file_menu.add_command(label="Save Progress", command=lambda: save_form_progress(window))
    unfinished_menu = tk.Menu(file_menu, tearoff=0)
    file_menu.add_cascade(label="Open Unfinished Reports", menu=unfinished_menu)
    unfinished_menu.configure(postcommand=lambda: rebuild_unfinished_menu(window, unfinished_menu))
    rebuild_unfinished_menu(window, unfinished_menu)
    file_menu.add_command(label="Manage Unfinished Reports...", command=lambda: open_unfinished_manager(window))
    file_menu.add_separator()
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
    tools_menu.add_separator()
    tools_menu.add_command(
        label="Report Number Prefix...",
        command=lambda: open_prefix_editor(window),
    )
    menubar.add_cascade(label="Tools", menu=tools_menu)

    help_menu = tk.Menu(menubar, tearoff=0)
    help_menu.add_command(label="How-To Guide...", command=lambda: open_howto(window))
    help_menu.add_command(label="Placeholder Index (PY_ tags)...", command=lambda: show_placeholder_guide(window))
    help_menu.add_separator()
    help_menu.add_command(label="About", command=lambda: show_about(window))
    menubar.add_cascade(label="Help", menu=help_menu)

    window.config(menu=menubar)
    return menubar


def save_form_progress(window):
    from drafts_manager import save_progress

    try:
        path, label = save_progress(window)
    except ValueError as exc:
        messagebox.showinfo("Save Progress", str(exc), parent=window)
        return
    except Exception as exc:
        messagebox.showerror("Save Progress", f"Could not save progress.\n\n{exc}", parent=window)
        return
    messagebox.showinfo("Save Progress", f"Saved unfinished report:\n{label}\n\n{path}", parent=window)


def _start_screen_of(window):
    if window is None:
        return None
    if window.__class__.__name__ == "StartScreen":
        return window
    master = getattr(window, "master", None)
    if master is not None and master.__class__.__name__ == "StartScreen":
        return master
    return None


def rebuild_unfinished_menu(window, menu):
    from drafts_manager import list_drafts

    menu.delete(0, "end")
    drafts = list_drafts()
    if not drafts:
        menu.add_command(label="(Empty)", state="disabled")
        return
    for item in drafts:
        caption = f"{item['label']}  ({item['report_label']})"
        menu.add_command(
            label=caption,
            command=lambda draft=item: open_unfinished_report(window, draft),
        )


def open_unfinished_report(window, item):
    from drafts_manager import apply_draft, launch_report_type, load_draft_file

    start = _start_screen_of(window)
    payload = dict(item.get("data") or {})
    payload["_path"] = str(item["path"])
    report_type = item.get("report_type")
    if start is None:
        if getattr(window, "report_type", None) != report_type:
            messagebox.showinfo(
                "Open Unfinished Reports",
                "Return to the start screen to open a different report type.",
                parent=window,
            )
            return
        try:
            apply_draft(window, payload)
        except Exception as exc:
            messagebox.showerror("Open Unfinished Reports", str(exc), parent=window)
        return

    open_app = getattr(start, "open_app", None)
    if open_app is not None:
        if not messagebox.askyesno(
            "Open Unfinished Reports",
            "Open this unfinished report and replace the form that is currently open?",
            parent=window,
        ):
            return
    try:
        data = load_draft_file(item["path"])
        data["_path"] = str(item["path"])
        launch_report_type(start, report_type, data)
    except Exception as exc:
        messagebox.showerror("Open Unfinished Reports", str(exc), parent=window)


def open_unfinished_manager(parent):
    from drafts_manager import delete_draft_file, list_drafts

    win = tk.Toplevel(parent)
    win.title("Manage Unfinished Reports")
    win.configure(bg=COLORS["bg"])
    win.geometry("720x460")
    win.minsize(560, 360)
    win.transient(parent)

    ttk.Label(
        win,
        text="Delete unfinished reports you no longer need. Generating a report also removes its draft.",
        wraplength=680,
        style="Hint.TLabel",
    ).pack(anchor="w", padx=12, pady=(10, 8))

    list_wrap = ttk.Frame(win)
    list_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 8))
    scroll = ttk.Scrollbar(list_wrap)
    scroll.pack(side="right", fill="y")
    listbox = tk.Listbox(
        list_wrap,
        bg=COLORS["entry_bg"],
        fg=COLORS["text"],
        selectbackground=COLORS["accent_dark"],
        highlightthickness=0,
        activestyle="none",
        exportselection=False,
        selectmode="extended",
    )
    listbox.pack(fill="both", expand=True)
    listbox.configure(yscrollcommand=scroll.set)
    scroll.configure(command=listbox.yview)

    rows = []

    def caption(item):
        updated = item.get("updated") or ""
        extra = f"  —  {updated}" if updated else ""
        return f"{item['label']}  ({item['report_label']}){extra}"

    def refresh():
        rows.clear()
        listbox.delete(0, "end")
        for item in list_drafts():
            rows.append(item)
            listbox.insert("end", caption(item))
        if not rows:
            listbox.insert("end", "(Empty)")
            listbox.itemconfig(0, fg=COLORS["muted"])

    def selected_items():
        if not rows:
            return []
        return [rows[index] for index in listbox.curselection() if 0 <= index < len(rows)]

    def clear_open_draft_if_removed(path):
        start = _start_screen_of(parent)
        targets = [parent]
        if start is not None:
            targets.append(start)
            open_app = getattr(start, "open_app", None)
            if open_app is not None:
                targets.append(open_app)
        removed = str(path)
        for widget in targets:
            stored = getattr(widget, "_draft_path", None)
            if stored and str(stored) == removed:
                widget._draft_path = None

    def delete_selected():
        items = selected_items()
        if not items:
            messagebox.showinfo("Manage Unfinished Reports", "Select an unfinished report first.", parent=win)
            return
        if len(items) == 1:
            prompt = f"Delete this unfinished report?\n\n{items[0]['label']}"
        else:
            prompt = f"Delete {len(items)} unfinished reports?"
        if not messagebox.askyesno("Delete Unfinished Report", prompt, parent=win):
            return
        failed = []
        for item in items:
            if delete_draft_file(item["path"]):
                clear_open_draft_if_removed(item["path"])
            else:
                failed.append(item["label"])
        refresh()
        if failed:
            messagebox.showerror(
                "Manage Unfinished Reports",
                "Could not delete:\n" + "\n".join(failed),
                parent=win,
            )

    def delete_all():
        if not rows:
            messagebox.showinfo("Manage Unfinished Reports", "There are no unfinished reports.", parent=win)
            return
        if not messagebox.askyesno(
            "Delete All Unfinished Reports",
            f"Delete all {len(rows)} unfinished reports?\n\nThis cannot be undone.",
            parent=win,
        ):
            return
        for item in list(rows):
            if delete_draft_file(item["path"]):
                clear_open_draft_if_removed(item["path"])
        refresh()

    actions = ttk.Frame(win)
    actions.pack(fill="x", padx=12, pady=(0, 12))
    ttk.Button(actions, text="Delete", command=delete_selected).pack(side="left")
    ttk.Button(actions, text="Delete All", command=delete_all).pack(side="left", padx=(8, 0))
    ttk.Button(actions, text="Close", command=win.destroy).pack(side="right")
    listbox.bind("<Delete>", lambda _event: delete_selected())
    refresh()


def open_prefix_editor(parent):
    from drafts_manager import apply_prefix_to_open_forms
    from report_common import current_dfr_prefix, save_dfr_prefix

    current = current_dfr_prefix()
    win = tk.Toplevel(parent)
    win.title("Report Number Prefix")
    win.configure(bg=COLORS["bg"])
    win.resizable(False, False)
    win.transient(parent)
    frame = ttk.Frame(win, padding=18)
    frame.pack(fill="both", expand=True)
    ttk.Label(
        frame,
        text="This prefix is filled into the Report Number field on new forms.",
        wraplength=420,
        style="Hint.TLabel",
    ).pack(anchor="w")
    ttk.Label(frame, text="Prefix:").pack(anchor="w", pady=(12, 4))
    entry = ttk.Entry(frame, width=36)
    entry.pack(anchor="w", fill="x")
    entry.insert(0, current)
    entry.focus_set()

    def save():
        try:
            new_prefix = save_dfr_prefix(entry.get())
        except ValueError as exc:
            messagebox.showerror("Report Number Prefix", str(exc), parent=win)
            return
        apply_prefix_to_open_forms(parent, current, new_prefix)
        win.destroy()

    buttons = ttk.Frame(frame)
    buttons.pack(fill="x", pady=(16, 0))
    ttk.Button(buttons, text="Save", command=save).pack(side="right")
    ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="right", padx=(0, 8))
    win.bind("<Return>", lambda _event: save())
    win.update_idletasks()
    width = max(420, win.winfo_reqwidth())
    height = win.winfo_reqheight()
    px = parent.winfo_rootx() + max(0, (parent.winfo_width() - width) // 2)
    py = parent.winfo_rooty() + max(0, (parent.winfo_height() - height) // 3)
    win.geometry(f"{width}x{height}+{px}+{py}")
    win.grab_set()


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
    root = window if is_start or master is None else master
    _autosave_window(window)
    _autosave_window(getattr(root, "open_app", None))
    _autosave_window(master)
    open_app = getattr(root, "open_app", None)
    if open_app is not None:
        try:
            root.open_app = None
        except Exception:
            pass
        try:
            open_app.protocol("WM_DELETE_WINDOW", lambda: None)
        except Exception:
            pass
        try:
            if open_app.winfo_exists():
                open_app.destroy()
        except Exception:
            pass
    if window is not root:
        try:
            window.protocol("WM_DELETE_WINDOW", lambda: None)
        except Exception:
            pass
        try:
            if window.winfo_exists():
                window.destroy()
        except Exception:
            pass
    try:
        if root is not None and root.winfo_exists():
            root.destroy()
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


def howto_path():
    names = (
        "Digital Forensics Report Writer How-To.docx",
        "How-To.docx",
    )
    roots = []
    try:
        roots.append(resource_dir())
    except Exception:
        pass
    try:
        roots.append(writable_dir())
    except Exception:
        pass
    roots.append(Path(__file__).resolve().parent)
    roots.append(Path(__file__).resolve().parent.parent)
    for folder in roots:
        for name in names:
            candidate = Path(folder) / name
            if candidate.is_file():
                return candidate
    return None


def open_howto(parent):
    path = howto_path()
    if path is None:
        messagebox.showinfo(
            "How-To Guide",
            "The How-To document was not found next to the program.\n"
            "Look for Digital Forensics Report Writer How-To.docx in the project folder.",
            parent=parent,
        )
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))
        else:
            import subprocess
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as exc:
        messagebox.showerror("How-To Guide", f"Could not open:\n{path}\n\n{exc}", parent=parent)


def show_placeholder_guide(parent):
    win = tk.Toplevel(parent)
    win.title("Placeholder Index")
    win.configure(bg=COLORS["bg"])
    win.geometry("780x640")
    ttk.Label(
        win,
        text=(
            "Use these names as typed PY_ tokens or as the Tag on a Word content control. "
            "Tags are preferred. Matching is the full name, so PY_EXAMINE does not fill "
            "PY_EXAMINEVER or check PY_EXAMINEYES. Official base templates are "
            "DFR Mobile.docx, DFR Computer.docx, DFR Storage.docx, and DFR SW Return.docx."
        ),
        wraplength=740,
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
