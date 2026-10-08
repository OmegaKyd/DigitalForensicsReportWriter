"""Tools → Edit Agencies window."""
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from report_common import (
    DEFAULT_AGENCIES,
    load_request_agencies,
    refresh_open_agency_comboboxes,
    revert_request_agencies,
    save_request_agencies,
)
from ui_theme import COLORS


def open_agencies_editor(parent):
    win = tk.Toplevel(parent)
    win.title("Edit Agencies")
    win.configure(bg=COLORS["bg"])
    win.geometry("520x460")
    win.minsize(420, 380)
    win.transient(parent)

    ttk.Label(
        win,
        text="These agencies are shared by Requesting Agency, Examiner Agency, and "
        "Transfer Officer Agency. Type an agency that is not listed; saving a report adds it "
        "here. Drag a row to reorder, or use Move Up / Move Down.",
        wraplength=490,
        style="Hint.TLabel",
    ).pack(anchor="w", padx=12, pady=(10, 6))

    body = ttk.Frame(win)
    body.pack(fill="both", expand=True, padx=12, pady=(0, 8))
    body.columnconfigure(0, weight=1)
    body.rowconfigure(0, weight=1)

    list_frame = ttk.Frame(body)
    list_frame.grid(row=0, column=0, sticky="nsew")
    list_frame.rowconfigure(0, weight=1)
    list_frame.columnconfigure(0, weight=1)

    scroll = ttk.Scrollbar(list_frame)
    scroll.grid(row=0, column=1, sticky="ns")
    title_list = tk.Listbox(
        list_frame,
        exportselection=False,
        bg=COLORS["entry_bg"],
        fg=COLORS["text"],
        selectbackground=COLORS["accent_dark"],
        highlightthickness=0,
        activestyle="none",
    )
    title_list.grid(row=0, column=0, sticky="nsew")
    title_list.configure(yscrollcommand=scroll.set)
    scroll.configure(command=title_list.yview)

    buttons = ttk.Frame(body)
    buttons.grid(row=0, column=1, sticky="ns", padx=(10, 0))

    drag = {"index": None}

    def current_titles():
        return list(title_list.get(0, "end"))

    def selected_index():
        selection = title_list.curselection()
        if not selection:
            return None
        return int(selection[0])

    def fill(titles, keep=""):
        title_list.delete(0, "end")
        for item in titles:
            title_list.insert("end", item)
        if keep:
            values = [item.casefold() for item in titles]
            if keep.casefold() in values:
                title_list.selection_set(values.index(keep.casefold()))
                title_list.see(values.index(keep.casefold()))
        elif titles:
            title_list.selection_set(0)

    def persist():
        saved = save_request_agencies(current_titles())
        refresh_open_agency_comboboxes(parent)
        return saved

    def add_title():
        value = simpledialog.askstring("Add Agency", "New agency name:", parent=win)
        if value is None:
            return
        text = " ".join(value.split())
        if not text:
            return
        existing = {item.casefold() for item in current_titles()}
        if text.casefold() in existing:
            messagebox.showinfo("Already listed", f'"{text}" is already in the list.', parent=win)
            return
        title_list.insert("end", text)
        title_list.selection_clear(0, "end")
        title_list.selection_set("end")
        title_list.see("end")
        persist()

    def rename_title():
        index = selected_index()
        if index is None:
            messagebox.showinfo("Select an agency", "Select an agency to rename.", parent=win)
            return
        current = title_list.get(index)
        value = simpledialog.askstring("Rename Agency", "Agency name:", initialvalue=current, parent=win)
        if value is None:
            return
        text = " ".join(value.split())
        if not text:
            return
        others = {item.casefold() for i, item in enumerate(current_titles()) if i != index}
        if text.casefold() in others:
            messagebox.showinfo("Already listed", f'"{text}" is already in the list.', parent=win)
            return
        title_list.delete(index)
        title_list.insert(index, text)
        title_list.selection_set(index)
        persist()

    def remove_title():
        index = selected_index()
        if index is None:
            messagebox.showinfo("Select an agency", "Select an agency to remove.", parent=win)
            return
        current = title_list.get(index)
        if not messagebox.askyesno("Remove Agency", f'Remove "{current}" from the list?', parent=win):
            return
        title_list.delete(index)
        if title_list.size() == 0:
            fill(DEFAULT_AGENCIES)
        elif index < title_list.size():
            title_list.selection_set(index)
        else:
            title_list.selection_set(title_list.size() - 1)
        persist()

    def move(delta):
        index = selected_index()
        if index is None:
            return
        target = index + delta
        if target < 0 or target >= title_list.size():
            return
        value = title_list.get(index)
        title_list.delete(index)
        title_list.insert(target, value)
        title_list.selection_set(target)
        title_list.see(target)
        persist()

    def on_press(event):
        drag["index"] = title_list.nearest(event.y)
        title_list.selection_clear(0, "end")
        title_list.selection_set(drag["index"])

    def on_drag(event):
        start = drag["index"]
        if start is None:
            return
        target = title_list.nearest(event.y)
        if target == start:
            return
        value = title_list.get(start)
        title_list.delete(start)
        title_list.insert(target, value)
        title_list.selection_clear(0, "end")
        title_list.selection_set(target)
        drag["index"] = target

    def on_release(_event):
        if drag["index"] is not None:
            persist()
        drag["index"] = None

    def revert():
        if not messagebox.askyesno(
            "Revert Agencies",
            "Restore the original agency list and discard custom agencies and order?",
            parent=win,
        ):
            return
        fill(revert_request_agencies())
        refresh_open_agency_comboboxes(parent)

    ttk.Button(buttons, text="Add...", command=add_title).pack(fill="x", pady=2)
    ttk.Button(buttons, text="Rename...", command=rename_title).pack(fill="x", pady=2)
    ttk.Button(buttons, text="Remove", command=remove_title).pack(fill="x", pady=2)
    ttk.Button(buttons, text="Move Up", command=lambda: move(-1)).pack(fill="x", pady=(12, 2))
    ttk.Button(buttons, text="Move Down", command=lambda: move(1)).pack(fill="x", pady=2)
    ttk.Button(buttons, text="Revert to Defaults", command=revert).pack(fill="x", pady=(12, 2))
    ttk.Button(buttons, text="Close", command=win.destroy).pack(fill="x", pady=(12, 2))

    title_list.bind("<ButtonPress-1>", on_press)
    title_list.bind("<B1-Motion>", on_drag)
    title_list.bind("<ButtonRelease-1>", on_release)
    title_list.bind("<Double-Button-1>", lambda _e: rename_title())
    title_list.bind("<Delete>", lambda _e: remove_title())

    fill(load_request_agencies())
    win.grab_set()
    win.focus_set()
    title_list.focus_set()
