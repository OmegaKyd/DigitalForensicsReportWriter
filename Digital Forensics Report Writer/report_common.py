"""Shared report helpers used by the Digital Forensics Report Writer modules.

Keeps PY_ placeholder replacement in the report windows. This module holds
classification, field merge, folder memory, output naming, validation extras,
and the pre-write preview so those behaviors stay in one place.
"""
import os
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

from settings_manager import SettingsManager
from ui_theme import COLORS, parse_mdy_date, resource_dir, writable_dir

TEMPLATES_FOLDER_NAME = "DFR Templates"
BUNDLED_TEMPLATES_NAME = "Templates"
TEMPLATES_DIR_KEY = "dfr_templates_dir"
REQUEST_TITLES_KEY = "remembered_request_titles"
REQUEST_AGENCIES_KEY = "remembered_request_agencies"
AGENCIES_SCHEMA_KEY = "request_agencies_schema"
AGENCIES_SCHEMA = 1
DEFAULT_AGENCIES = [
    "South Dakota DCI",
    "South Dakota Highway Patrol",
]
AGENCY_COMBO_ATTRS = (
    "request_agency",
    "requesting_agency",
    "transfer_agency",
    "examiner_agency_type",
    "examiner_agency",
)
DEFAULT_REQUEST_TITLES = [
    "Officer",
    "Deputy",
    "Detective",
    "Investigator",
    "Trooper",
    "Sergeant",
    "Lieutenant",
    "Captain",
    "Special Agent",
    "Supervisory Special Agent",
    "SAAG",
    "Intel Analyst",
    "Computer Forensic Examiner",
    "Digital Forensic Examiner",
    "Forensic Analyst",
]
TITLES_SCHEMA_KEY = "request_titles_schema"
TITLES_SCHEMA = 2


def _clean_label(value):
    return " ".join(str(value or "").split())


def title_agency_words(agency):
    """Title-case an agency without flattening capitals the user already typed."""
    parts = []
    for word in str(agency or "").split():
        if word[0].isupper() or any(ch.isupper() for ch in word[1:]):
            parts.append(word)
        else:
            parts.append(word.capitalize())
    return " ".join(parts)


def _normalize_titles(values):
    titles = []
    seen = set()
    for value in values or []:
        text = _clean_label(value)
        key = text.casefold()
        if text and key not in {"other", "other (specify)"} and key not in seen:
            seen.add(key)
            titles.append(text)
    return titles


def load_request_titles():
    settings = SettingsManager().load_settings()
    remembered = _normalize_titles(settings.get(REQUEST_TITLES_KEY) or [])
    if settings.get(TITLES_SCHEMA_KEY) == TITLES_SCHEMA:
        return remembered or list(DEFAULT_REQUEST_TITLES)
    return _normalize_titles(list(remembered) + list(DEFAULT_REQUEST_TITLES))


def save_request_titles(titles):
    cleaned = _normalize_titles(titles)
    if not cleaned:
        cleaned = list(DEFAULT_REQUEST_TITLES)
    SettingsManager().save_settings({REQUEST_TITLES_KEY: cleaned, TITLES_SCHEMA_KEY: TITLES_SCHEMA})
    return cleaned


def revert_request_titles():
    SettingsManager().save_settings({REQUEST_TITLES_KEY: [], TITLES_SCHEMA_KEY: TITLES_SCHEMA})
    return list(DEFAULT_REQUEST_TITLES)


def remember_request_title(title):
    text = _clean_label(title)
    titles = load_request_titles()
    if not text or text.casefold() in {"other", "other (specify)"}:
        return titles
    if text.casefold() not in {item.casefold() for item in titles}:
        titles.append(text)
        save_request_titles(titles)
    return load_request_titles()


def refresh_open_title_comboboxes(root):
    titles = load_request_titles()
    if root is None:
        return titles
    seen = set()
    stack = [root]
    master = getattr(root, "master", None)
    if master is not None:
        stack.append(master)
    while stack:
        widget = stack.pop()
        ident = id(widget)
        if ident in seen:
            continue
        seen.add(ident)
        for attr in ("request_title_type", "requesting_officer_title", "examiner_title_type", "examiner_title", "transfer_title"):
            box = getattr(widget, attr, None)
            if box is None:
                continue
            try:
                current = box.get()
                box.configure(values=titles)
                if current:
                    box.set(current)
            except Exception:
                pass
        try:
            stack.extend(widget.winfo_children())
        except Exception:
            pass
    return titles


def bind_prefix_typeahead(combobox):
    """Suggest a list match after the typed prefix.

    The characters already typed are never rewritten. If nothing in the
    list starts with that prefix, the field is left exactly as typed so
    capitals such as the C in City are kept.
    """
    state = {"prefix": "", "job": None}

    def options():
        return [str(item) for item in (combobox.cget("values") or ())]

    def first_match(prefix):
        if not prefix:
            return None
        needle = prefix.casefold()
        for item in options():
            if item.casefold().startswith(needle):
                return item
        return None

    def apply(prefix):
        state["job"] = None
        match = first_match(prefix)
        if not match:
            return
        text = prefix + match[len(prefix):]
        combobox.set(text)
        try:
            combobox.icursor(len(prefix))
            if len(text) > len(prefix):
                combobox.selection_range(len(prefix), len(text))
        except Exception:
            pass

    def schedule(prefix):
        if state["job"] is not None:
            try:
                combobox.after_cancel(state["job"])
            except Exception:
                pass
        state["job"] = combobox.after_idle(lambda p=prefix: apply(p))

    def commit(_event=None):
        if state["job"] is not None:
            try:
                combobox.after_cancel(state["job"])
            except Exception:
                pass
            state["job"] = None
        current = combobox.get() or ""
        stripped = current.strip()
        exact = None
        needle = stripped.casefold()
        if needle:
            for item in options():
                if item.casefold() == needle:
                    exact = item
                    break
        if exact:
            combobox.set(exact)
            state["prefix"] = exact
        else:
            state["prefix"] = current
        try:
            combobox.selection_clear()
            combobox.icursor("end")
        except Exception:
            pass

    def on_key(event):
        if event.keysym in ("Tab", "Return"):
            commit()
            return
        if event.keysym in ("Up", "Down", "Left", "Right", "Escape"):
            return
        if event.keysym in ("BackSpace", "Delete"):
            state["prefix"] = combobox.get()
            return
        char = event.char or ""
        if len(char) != 1 or not char.isprintable():
            return
        prev = state["prefix"]
        current = combobox.get()
        completed = first_match(prev) if prev else None
        if completed and current.casefold().startswith((prev + char).casefold()):
            prefix = prev + char
        elif current.casefold().endswith(char.casefold()):
            prefix = current
        else:
            prefix = prev + char
        state["prefix"] = prefix
        schedule(prefix)

    def on_select(_event=None):
        state["prefix"] = combobox.get()
        try:
            combobox.selection_clear()
        except Exception:
            pass

    combobox.bind("<KeyRelease>", on_key, add="+")
    combobox.bind("<<ComboboxSelected>>", on_select, add="+")
    combobox.bind("<FocusOut>", commit, add="+")


def setup_title_combobox(widget, saved="", on_change=None):
    saved_text = _clean_label(saved)
    titles = load_request_titles()
    if saved_text:
        titles = remember_request_title(saved_text)
    widget.configure(values=titles)
    bind_prefix_typeahead(widget)
    widget.set(saved_text)
    if on_change is not None:
        widget.bind("<<ComboboxSelected>>", on_change, add="+")
        widget.bind("<FocusOut>", on_change, add="+")
    return widget


def remember_titles_from_form(app):
    fields = []
    for attr in ("request_title_type", "requesting_officer_title", "examiner_title_type", "examiner_title", "transfer_title"):
        widget = getattr(app, attr, None)
        if widget is None:
            continue
        try:
            value = widget.get()
        except Exception:
            value = ""
        remember_request_title(value)
        fields.append(widget)
    titles = load_request_titles()
    for widget in fields:
        try:
            current = widget.get()
            widget.configure(values=titles)
            if current:
                widget.set(current)
        except Exception:
            pass
    refresh_open_title_comboboxes(app)
    return titles


def saved_examiner_agency(settings=None):
    """Last typed/selected examiner agency, resolving the old Other (specify) pair."""
    data = settings if isinstance(settings, dict) else _settings().load_settings()
    kind = _clean_label(data.get("examiner_agency_type", ""))
    custom = _clean_label(data.get("examiner_agency_custom", ""))
    if kind.casefold() in {"other", "other (specify)"}:
        return custom
    return kind


def _normalize_agencies(values):
    return _normalize_titles(values)


def load_request_agencies():
    settings = _settings().load_settings()
    remembered = _normalize_agencies(settings.get(REQUEST_AGENCIES_KEY) or [])
    seeded = list(remembered)
    last = saved_examiner_agency(settings)
    if last:
        seeded.append(last)
    if settings.get(AGENCIES_SCHEMA_KEY) == AGENCIES_SCHEMA:
        return _normalize_agencies(seeded) or list(DEFAULT_AGENCIES)
    return _normalize_agencies(list(seeded) + list(DEFAULT_AGENCIES))


def save_request_agencies(agencies):
    cleaned = _normalize_agencies(agencies)
    if not cleaned:
        cleaned = list(DEFAULT_AGENCIES)
    _settings().save_settings({REQUEST_AGENCIES_KEY: cleaned, AGENCIES_SCHEMA_KEY: AGENCIES_SCHEMA})
    return cleaned


def revert_request_agencies():
    _settings().save_settings({REQUEST_AGENCIES_KEY: [], AGENCIES_SCHEMA_KEY: AGENCIES_SCHEMA})
    return list(DEFAULT_AGENCIES)


def remember_request_agency(agency):
    text = _clean_label(agency)
    agencies = load_request_agencies()
    if not text or text.casefold() in {"other", "other (specify)"}:
        return agencies
    if text.casefold() not in {item.casefold() for item in agencies}:
        agencies.append(text)
        save_request_agencies(agencies)
    return load_request_agencies()


def refresh_open_agency_comboboxes(root):
    agencies = load_request_agencies()
    if root is None:
        return agencies
    seen = set()
    stack = [root]
    master = getattr(root, "master", None)
    if master is not None:
        stack.append(master)
    while stack:
        widget = stack.pop()
        ident = id(widget)
        if ident in seen:
            continue
        seen.add(ident)
        for attr in AGENCY_COMBO_ATTRS:
            box = getattr(widget, attr, None)
            if box is None:
                continue
            try:
                current = box.get()
                box.configure(values=agencies)
                if current:
                    box.set(current)
            except Exception:
                pass
        try:
            stack.extend(widget.winfo_children())
        except Exception:
            pass
    return agencies


def setup_agency_combobox(widget, saved="", on_change=None):
    saved_text = _clean_label(saved)
    agencies = load_request_agencies()
    if saved_text:
        agencies = remember_request_agency(saved_text)
    widget.configure(values=agencies)
    bind_prefix_typeahead(widget)
    widget.set(saved_text)
    if on_change is not None:
        widget.bind("<<ComboboxSelected>>", on_change, add="+")
        widget.bind("<FocusOut>", on_change, add="+")
    return widget


def remember_agencies_from_form(app):
    fields = []
    for attr in AGENCY_COMBO_ATTRS:
        widget = getattr(app, attr, None)
        if widget is None:
            continue
        try:
            value = widget.get()
        except Exception:
            value = ""
        remember_request_agency(value)
        fields.append(widget)
    agencies = load_request_agencies()
    for widget in fields:
        try:
            current = widget.get()
            widget.configure(values=agencies)
            if current:
                widget.set(current)
        except Exception:
            pass
    refresh_open_agency_comboboxes(app)
    return agencies


def refresh_request_title_values(widget, title=""):
    titles = remember_request_title(title) if title else load_request_titles()
    if widget is None:
        return titles
    try:
        current = widget.get()
        widget.configure(values=titles)
        if current:
            widget.set(current)
    except Exception:
        pass
    return titles

# First-listed rule for identifier lists in Cellebrite Quick View / similar reports.
FIRST_LISTED_FIELDS = ("device_iccid", "DEV_IMEI", "device_carrier")


def looks_like_digital_collector_log(content):
    text = (content or "").lstrip("\ufeff")
    return bool(re.search(r"Digital Collector Version\s*:", text, re.I))


def parse_digital_collector_date(date_str):
    raw = (date_str or "").strip()
    if not raw:
        return ""
    tz = ""
    tz_match = re.search(r"\(([^)]+)\)", raw)
    if tz_match:
        tz = tz_match.group(1).strip()
    cleaned = re.sub(r"\s*\([^)]*\)", "", raw).strip()
    try:
        date_obj = datetime.strptime(cleaned, "%Y-%m-%d %H:%M:%S")
        formatted = date_obj.strftime("%A, %B %d, %Y at %H:%M")
        formatted = formatted.replace(" 0", " ")
        return f"{formatted} {tz}".strip()
    except Exception:
        return raw


def parse_digital_collector_log(content):
    """Parse a Cellebrite Digital Collector acquisition log.

    Capacity uses the volume GB value (e.g. 118.0 GB) when a Device Identifier
    block is present; otherwise the first disk GB value.
    """
    text = (content or "").lstrip("\ufeff")
    data = {
        "extraction_tool": "Cellebrite Digital Collector",
        "extraction_type": "Disk Imaging",
    }
    populated = set()

    def grab(pattern, flags=re.I | re.M):
        match = re.search(pattern, text, flags)
        if not match:
            return ""
        return match.group(1).strip()

    version = grab(r"Digital Collector Version[ \t]*:[ \t]*(.+)")
    if version:
        data["DC_OS"] = version
        populated.add("DC_OS")

    mapping = [
        (r"Case Name[ \t]*:[ \t]*(.*)$", "case_name"),
        (r"Case Number/ID[ \t]*:[ \t]*(.*)$", "case_number"),
        (r"Location[ \t]*:[ \t]*(.*)$", "location"),
        (r"Exhibit ID/Evidence #[ \t]*:[ \t]*(.*)$", "evidence_number"),
        (r"Description[ \t]*:[ \t]*(.*)$", "case_notes"),
        (r"Examiner[ \t]*:[ \t]*(.*)$", "examiner"),
        (r"Agency/Company[ \t]*:[ \t]*(.*)$", "examiner_agency"),
        (r"Source Device[ \t]*:[ \t]*(.*)$", "source_path"),
        (r"Model[ \t]*:[ \t]*(.*)$", "device_model"),
        (r"Serial Number[ \t]*:[ \t]*(.*)$", "device_serial"),
        (r"Format[ \t]*:[ \t]*(.*)$", "image_format"),
        (r"md5[ \t]*:[ \t]*([A-Fa-f0-9]+)", "md5_hash"),
        (r"sha1[ \t]*:[ \t]*([A-Fa-f0-9]+)", "sha1_hash"),
    ]
    for pattern, key in mapping:
        value = grab(pattern)
        if value:
            data[key] = value
            populated.add(key)

    volume_capacity = ""
    volume_block = re.search(
        r"Device Identifier\s*:.*?(?:\n\s*\n|\nDestination |\Z)",
        text,
        re.I | re.S,
    )
    cap_re = re.compile(r"Capacity\s*:\s*(\d+(?:\.\d+)?)\s*GB\b", re.I)
    if volume_block:
        match = cap_re.search(volume_block.group(0))
        if match:
            volume_capacity = f"{match.group(1)} GB"
    if not volume_capacity:
        match = cap_re.search(text)
        if match:
            volume_capacity = f"{match.group(1)} GB"
    if volume_capacity:
        data["device_capacity"] = volume_capacity
        populated.add("device_capacity")

    start = grab(r"Acquisition Start Time[ \t]*:[ \t]*(.+)")
    if start:
        data["extraction_date"] = start
        data["formatted_date"] = parse_digital_collector_date(start)
        populated.update({"extraction_date", "formatted_date"})
    else:
        data["formatted_date"] = "Unknown Date"
        populated.add("formatted_date")

    return data, populated

GUI_PROTECTED_FIELDS = (
    "device_iccid",
    "device_passcode",
    "device_carrier",
    "device_color",
    "device_capacity",
    "Device_Color",
    "Device_Capacity",
    "device_password",
    "Phone_Number",
    "DEV_IMEI",
    "Serial_Number",
    "Device_Name",
    "Device_Account",
    "Case_Number",
    "evidence_ID",
    "DFR_Num",
    "Device_Owner",
    "Examiner_Name",
    "Examiner_Title",
    "Examiner_Agency",
    "Request_Officer",
    "Request_Agency",
    "Request_Title",
    "hd_make",
    "hd_model",
    "hd_serial",
    "device_PCMan",
    "device_PCMod",
    "device_PCSerial",
)

def current_dfr_prefix():
    """Prefix for a new report number, based on the year the program is opened."""
    return f"DFR{datetime.now().year}-"


def is_complete_dfr_number(value):
    """False when empty or still only the DFRYYYY- prefix."""
    text = (value or "").strip()
    if not text:
        return False
    return not bool(re.fullmatch(r"DFR\d{4}-?", text, re.I))


FOLDER_KEYS = {
    "template": "last_template_dir",
    "extraction": "last_extraction_dir",
    "export": "last_export_dir",
}

PDF_KIND_LABELS = {
    "graykey": "GrayKey Progress Report",
    "summary": "Cellebrite Summary Report",
    "quickview": "Cellebrite Quick View",
    "unknown": "Unrecognized PDF",
}


def bundled_templates_dir():
    """Official templates packed with the program. Do not write here when frozen."""
    return resource_dir() / BUNDLED_TEMPLATES_NAME


def _as_templates_folder(path):
    """Use the chosen folder if it already is DFR Templates (or legacy Templates)."""
    folder = Path(path)
    name = folder.name.lower()
    if name in (TEMPLATES_FOLDER_NAME.lower(), "templates"):
        return folder
    return folder / TEMPLATES_FOLDER_NAME


def _saved_templates_dir():
    raw = (_settings().load_settings().get(TEMPLATES_DIR_KEY) or "").strip()
    if raw and Path(raw).is_dir():
        return Path(raw)
    return None


def _remember_templates_dir(path):
    if not path:
        return
    manager = _settings()
    settings = manager.load_settings()
    settings[TEMPLATES_DIR_KEY] = str(path)
    manager.save_settings(settings)


def seed_templates(dest):
    """Copy missing official .docx files into the user DFR Templates folder."""
    dest = Path(dest)
    dest.mkdir(parents=True, exist_ok=True)
    bundled = bundled_templates_dir()
    try:
        if bundled.is_dir() and bundled.resolve() != dest.resolve():
            for src in bundled.glob("*.docx"):
                if src.name.startswith("~$"):
                    continue
                target = dest / src.name
                if not target.exists():
                    shutil.copy2(src, target)
    except Exception:
        pass
    return dest


def _folder_beside_exe(name):
    candidate = writable_dir() / name
    return candidate if candidate.is_dir() else None


def templates_dir():
    """Writable DFR Templates folder chosen by the user; seeded from the bundle."""
    dest = _saved_templates_dir()
    if dest is None:
        dest = _folder_beside_exe(TEMPLATES_FOLDER_NAME)
    if dest is None and not getattr(sys, "frozen", False):
        bundled = bundled_templates_dir()
        if bundled.is_dir():
            dest = bundled
    if dest is None:
        dest = writable_dir() / TEMPLATES_FOLDER_NAME
    return seed_templates(dest)


def _pick_templates_parent(parent, initial=None):
    from tkinter import filedialog, messagebox

    if parent is not None:
        try:
            messagebox.showinfo(
                "DFR Templates Folder",
                "Choose where the DFR Templates folder should be stored.\n\n"
                "A folder named \"DFR Templates\" will be created in that location "
                "and the official templates will be copied there.\n\n"
                "You can move it later with Tools → Change Template Folder Location.",
                parent=parent,
            )
        except Exception:
            pass
        chosen = filedialog.askdirectory(
            parent=parent,
            title="Choose location for the DFR Templates folder",
            initialdir=str(initial or writable_dir()),
        )
    else:
        chosen = ""
    if chosen:
        return _as_templates_folder(chosen)
    return writable_dir() / TEMPLATES_FOLDER_NAME


def _ask_about_found_templates(parent, found):
    """Ask before treating a generic Templates folder next to the exe as ours."""
    from tkinter import messagebox

    is_ours = messagebox.askyesno(
        "Found a Templates Folder",
        f"A folder named \"Templates\" was found next to the program:\n\n{found}\n\n"
        "Is this the DFR Templates folder used by this program?\n\n"
        "Choose No if this folder belongs to something else. "
        "It will be left alone and you can pick a new location.",
        parent=parent,
    )
    if not is_ours:
        return _pick_templates_parent(parent, initial=found.parent)

    renamed = found.parent / TEMPLATES_FOLDER_NAME
    if renamed.exists() and renamed.resolve() != found.resolve():
        use_existing = messagebox.askyesno(
            "DFR Templates Folder",
            f"A folder named \"{TEMPLATES_FOLDER_NAME}\" already exists next to the program.\n\n"
            "Use that folder instead? Choose No to pick a different location.",
            parent=parent,
        )
        if use_existing:
            return renamed
        return _pick_templates_parent(parent, initial=found.parent)

    rename_it = messagebox.askyesno(
        "Rename Templates Folder",
        "Would you like to rename this folder to \"DFR Templates\"?\n\n"
        "Choose Yes to rename it in place.\n"
        "Choose No to pick a new location instead. "
        "The existing Templates folder will be left as it is.",
        parent=parent,
    )
    if rename_it:
        try:
            found.rename(renamed)
            return renamed
        except Exception as exc:
            messagebox.showerror(
                "DFR Templates",
                f"Could not rename the folder:\n{exc}\n\n"
                "Choose a new location instead.",
                parent=parent,
            )
            return _pick_templates_parent(parent, initial=found.parent)
    return _pick_templates_parent(parent, initial=found.parent)


def ensure_user_templates(parent=None):
    """First run: ask where to create DFR Templates, then remember that path."""
    saved = _saved_templates_dir()
    if saved is not None:
        return seed_templates(saved)

    existing_dfr = _folder_beside_exe(TEMPLATES_FOLDER_NAME)
    if existing_dfr is not None:
        _remember_templates_dir(existing_dfr)
        return seed_templates(existing_dfr)

    found_templates = _folder_beside_exe("Templates")
    if found_templates is not None and getattr(sys, "frozen", False):
        dest = _ask_about_found_templates(parent, found_templates)
        seed_templates(dest)
        _remember_templates_dir(dest)
        return dest

    if not getattr(sys, "frozen", False):
        bundled = bundled_templates_dir()
        if bundled.is_dir():
            _remember_templates_dir(bundled)
            return bundled

    dest = _pick_templates_parent(parent)
    seed_templates(dest)
    _remember_templates_dir(dest)
    return dest


def change_templates_location(parent=None):
    """Move the DFR Templates folder to a new parent directory."""
    from tkinter import filedialog, messagebox

    old = templates_dir()
    chosen = filedialog.askdirectory(
        parent=parent,
        title="Choose a new location for the DFR Templates folder",
        initialdir=str(old.parent),
    )
    if not chosen:
        return None
    new = _as_templates_folder(chosen)
    try:
        if new.resolve() == old.resolve():
            messagebox.showinfo(
                "DFR Templates",
                "That is already the current templates folder.",
                parent=parent,
            )
            return old
    except Exception:
        pass

    try:
        new.parent.mkdir(parents=True, exist_ok=True)
        bundled = bundled_templates_dir()
        moving_official = False
        try:
            moving_official = old.resolve() == bundled.resolve()
        except Exception:
            moving_official = False

        if moving_official:
            # Never relocate the official project/bundled Templates folder.
            new.mkdir(parents=True, exist_ok=True)
            for item in old.iterdir():
                if item.name.startswith("~$"):
                    continue
                target = new / item.name
                if not target.exists():
                    if item.is_dir():
                        shutil.copytree(item, target)
                    else:
                        shutil.copy2(item, target)
        elif new.exists():
            for item in old.iterdir():
                target = new / item.name
                if target.exists():
                    continue
                shutil.move(str(item), str(target))
            try:
                if not any(old.iterdir()):
                    old.rmdir()
            except Exception:
                pass
        else:
            shutil.move(str(old), str(new))
        seed_templates(new)
        _remember_templates_dir(new)
        messagebox.showinfo(
            "DFR Templates",
            f"Templates folder moved to:\n{new}",
            parent=parent,
        )
        return new
    except Exception as exc:
        messagebox.showerror(
            "DFR Templates",
            f"Could not move the templates folder:\n{exc}",
            parent=parent,
        )
        return None


def list_dfr_templates(keywords=None):
    folder = templates_dir()
    if not folder.is_dir():
        return []
    items = []
    keys = [str(key).lower() for key in (keywords or ()) if str(key).strip()]
    for path in sorted(folder.glob("*.docx")):
        if path.name.startswith("~$"):
            continue
        filename = path.name.lower()
        if keys and not any(key in filename for key in keys):
            continue
        items.append({"name": path.stem, "path": str(path)})
    return items


def default_template_for(preferred="", keywords=None):
    """Pick the preferred file when it is in the filtered list; otherwise the first match."""
    items = list_dfr_templates(keywords)
    if not items:
        return None
    wanted = (preferred or "").strip()
    if wanted:
        wanted_name = Path(wanted).name.lower()
        wanted_stem = Path(wanted).stem.lower()
        for item in items:
            path_name = Path(item["path"]).name.lower()
            if path_name == wanted_name or item["name"].lower() == wanted_stem:
                return item
    return items[0]


def sync_template_choice(app, path):
    if not path:
        return
    name = Path(path).stem
    choices = getattr(app, "_template_choices", None)
    if choices is None:
        app._template_choices = {}
        choices = app._template_choices
    choices[name] = path
    combo = getattr(app, "template_choice", None)
    if combo is None:
        return
    try:
        combo.configure(values=list(choices.keys()))
        combo.set(name)
    except Exception:
        pass


def add_template_picker(app, parent, preferred="", keywords=None):
    """Dropdown of Templates/*.docx matching keywords. preferred is the default if present."""
    import tkinter as tk
    from tkinter import ttk

    items = list_dfr_templates(keywords)
    app._template_keywords = keywords
    app._template_preferred = preferred
    app._template_choices = {item["name"]: item["path"] for item in items}

    row = ttk.Frame(parent)
    row.pack(fill=tk.X, padx=5, pady=(6, 2))
    ttk.Label(row, text="DFR Template:").pack(side=tk.LEFT)
    app.template_choice = ttk.Combobox(
        row,
        state="readonly",
        values=list(app._template_choices.keys()),
    )
    app.template_choice.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

    def on_select(_event=None):
        name = app.template_choice.get()
        path = app._template_choices.get(name)
        if not path:
            return
        app.template_file = path
        for attr in ("template_drop_label", "template_status_label"):
            label = getattr(app, attr, None)
            if label is not None:
                try:
                    label.configure(text=f"Selected: {os.path.basename(path)}")
                except Exception:
                    pass

    app.template_choice.bind("<<ComboboxSelected>>", on_select)
    picked = default_template_for(preferred, keywords)
    if picked:
        app.template_choice.set(picked["name"])
        on_select()
    return app.template_choice


def first_listed(value):
    """Keep only the first identifier when a field holds a comma/bullet list."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    for sep in ("\n", ";", "|"):
        text = text.replace(sep, ",")
    parts = [part.strip(" •-\t") for part in text.split(",") if part.strip(" •-\t")]
    return parts[0] if parts else ""


def apply_first_listed_rule(data):
    if not data:
        return data
    for key in FIRST_LISTED_FIELDS:
        if data.get(key):
            data[key] = first_listed(data[key])
    return data


def prefer_gui_over_parsed(gui_data, parsed_data):
    """Keep examiner-typed values; fill blanks from the parsed extraction file."""
    merged = dict(parsed_data or {})
    for key, value in (gui_data or {}).items():
        if value not in (None, "") or key not in merged:
            merged[key] = value
    apply_first_listed_rule(merged)
    return merged


def _fill_entry_if_blank(widget, value):
    if widget is None or not value:
        return False
    try:
        current = (widget.get() or "").strip()
    except Exception:
        return False
    if current:
        return False
    widget.delete(0, "end")
    widget.insert(0, value)
    return True


def apply_log_device_fields_to_form(app, data):
    """Fill blank device fields from a parsed acquisition log. Typed values stay."""
    data = data or {}
    model = (data.get("device_model") or "").strip()
    serial = (data.get("device_serial") or "").strip()
    capacity = (data.get("device_capacity") or "").strip()
    device_type = ""
    combo = getattr(app, "device_type", None)
    if combo is not None:
        try:
            device_type = (combo.get() or "").strip()
        except Exception:
            device_type = ""
    if device_type == "Other Storage Device":
        custom = getattr(app, "device_type_entry", None)
        if custom is not None:
            try:
                typed = (custom.get() or "").strip()
            except Exception:
                typed = ""
            if typed:
                device_type = typed

    _fill_entry_if_blank(getattr(app, "device_capacity", None), capacity)
    if device_type == "Computer":
        _fill_entry_if_blank(getattr(app, "hd_model", None), model)
        _fill_entry_if_blank(getattr(app, "hd_serial", None), serial)
    else:
        _fill_entry_if_blank(getattr(app, "device_PCMod", None), model)
        _fill_entry_if_blank(getattr(app, "device_PCSerial", None), serial)


def merge_log_device_into_report_data(data, extraction_data, device_type=""):
    """Use log model/serial/capacity for tokens when the form left those fields blank."""
    data = data or {}
    extraction_data = extraction_data or {}
    model = (extraction_data.get("device_model") or "").strip()
    serial = (extraction_data.get("device_serial") or "").strip()
    capacity = (extraction_data.get("device_capacity") or "").strip()
    kind = (device_type or data.get("Device_Type") or data.get("device_type") or "").strip()
    if kind.lower() == "computer":
        if not (data.get("hd_model") or "").strip() and model:
            data["hd_model"] = model
        if not (data.get("hd_serial") or "").strip() and serial:
            data["hd_serial"] = serial
    else:
        if not (data.get("device_PCMod") or data.get("Device_Model") or "").strip() and model:
            data["device_PCMod"] = model
            data["Device_Model"] = model
        if not (data.get("device_PCSerial") or data.get("Device_Serial") or "").strip() and serial:
            data["device_PCSerial"] = serial
            data["Device_Serial"] = serial
    if not (data.get("Device_Capacity") or data.get("device_capacity") or "").strip() and capacity:
        data["Device_Capacity"] = capacity
        data["device_capacity"] = capacity
    return data


def _settings():
    return SettingsManager()


def get_last_folder(kind):
    key = FOLDER_KEYS.get(kind, kind)
    settings = _settings().load_settings()
    path = settings.get(key) or ""
    if path and os.path.isdir(path):
        return path
    return ""


def remember_folder(kind, path):
    if not path:
        return
    folder = path if os.path.isdir(path) else os.path.dirname(path)
    if not folder or not os.path.isdir(folder):
        return
    manager = _settings()
    settings = manager.load_settings()
    settings[FOLDER_KEYS.get(kind, kind)] = folder
    manager.save_settings(settings)


def ask_open_file(filetypes, folder_kind="extraction", title="Select file"):
    from tkinter import filedialog
    initial = get_last_folder(folder_kind)
    kwargs = {"filetypes": filetypes, "title": title}
    if initial:
        kwargs["initialdir"] = initial
    path = filedialog.askopenfilename(**kwargs)
    if path:
        remember_folder(folder_kind, path)
    return path


def ask_open_files(filetypes, folder_kind="extraction", title="Select file"):
    from tkinter import filedialog
    initial = get_last_folder(folder_kind)
    kwargs = {"filetypes": filetypes, "title": title}
    if initial:
        kwargs["initialdir"] = initial
    paths = filedialog.askopenfilenames(**kwargs)
    if paths:
        remember_folder(folder_kind, paths[0])
    return list(paths)


def ask_directory(folder_kind="export", current="", title="Select folder"):
    from tkinter import filedialog
    initial = current or get_last_folder(folder_kind)
    kwargs = {"title": title}
    if initial:
        kwargs["initialdir"] = initial
    path = filedialog.askdirectory(**kwargs)
    if path:
        remember_folder(folder_kind, path)
    return path


def default_export_dir():
    remembered = get_last_folder("export")
    if remembered:
        return remembered
    return os.path.join(os.path.expanduser("~"), "Desktop")


def sanitize_filename(name):
    cleaned = re.sub(r'[<>:"/\\|?*]', "", name or "").strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned or "DFR_Report"


def suggested_report_filename(dfr_number, module_label, owner="", model=""):
    """Default name: '(DFR #) - (Owner Name) (Device Model).docx'."""
    dfr = sanitize_filename(dfr_number or "")
    owner = sanitize_filename((owner or "").title())
    model = sanitize_filename(model or "")
    owner_model = " ".join(part for part in (owner, model) if part).strip()
    if dfr and owner_model:
        base = f"{dfr} - {owner_model}"
    elif dfr:
        base = dfr
    elif owner_model:
        base = owner_model
    else:
        base = "DFR_Report"
    return f"{base}.docx"


def unique_output_path(folder, filename):
    filename = sanitize_filename(os.path.splitext(filename)[0]) + ".docx"
    folder = folder or default_export_dir()
    path = os.path.join(folder, filename)
    stem, ext = os.path.splitext(filename)
    counter = 1
    while os.path.exists(path):
        path = os.path.join(folder, f"{stem} ({counter}){ext}")
        counter += 1
    return path


def classify_log_file(path, text=""):
    name = os.path.basename(path or "").lower()
    body = text or ""
    if not body and path and os.path.isfile(path) and path.lower().endswith(".txt"):
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as handle:
                body = handle.read(4000)
        except Exception:
            body = ""
    if "TX1 Log Entry" in body:
        return "tx1"
    if "FTK" in body and "Imager" in body:
        return "ftk"
    if "X-Ways Forensics" in body:
        return "xways"
    if name.endswith(".ufd"):
        return "ufd"
    return "unknown"


def confirm_pdf_kind(parent, guessed_kind, filename):
    """Trust a confident match; ask only when the PDF type is unclear."""
    from tkinter import messagebox, Toplevel, StringVar
    from tkinter import ttk
    import tkinter as tk

    if guessed_kind in ("graykey", "summary", "quickview"):
        return guessed_kind

    choice = {"kind": guessed_kind or "unknown"}

    dialog = Toplevel(parent)
    dialog.title("Identify PDF")
    dialog.transient(parent)
    dialog.grab_set()
    dialog.configure(bg=COLORS.get("bg", "#0b1e30"))

    ttk.Label(
        dialog,
        text=f"{os.path.basename(filename)} was not recognized automatically.\nSelect the report type:",
        justify="left",
    ).pack(padx=16, pady=(14, 8))

    selected = StringVar(value="graykey")
    for kind, label in (
        ("graykey", "GrayKey Progress Report"),
        ("summary", "Cellebrite Summary Report"),
        ("quickview", "Cellebrite Quick View"),
    ):
        ttk.Radiobutton(dialog, text=label, variable=selected, value=kind).pack(anchor="w", padx=24)

    def accept():
        choice["kind"] = selected.get()
        dialog.destroy()

    def cancel():
        choice["kind"] = "unknown"
        dialog.destroy()

    buttons = ttk.Frame(dialog)
    buttons.pack(fill="x", padx=16, pady=12)
    ttk.Button(buttons, text="Use this type", command=accept).pack(side="right", padx=4)
    ttk.Button(buttons, text="Cancel", command=cancel).pack(side="right")

    parent.wait_window(dialog)
    return choice["kind"]


def require_core_files(app, extraction_label="Extraction File", template_label="Template File"):
    missing = []
    if not getattr(app, "extraction_file", None):
        missing.append(extraction_label)
    if not getattr(app, "template_file", None):
        missing.append(template_label)
    return missing


def require_device_identity(extraction_data, manufacturer_keys=None, model_keys=None):
    manufacturer_keys = manufacturer_keys or ("device_manufacturer", "device_PCMan", "hd_make")
    model_keys = model_keys or ("device_model", "device_PCMod", "hd_model")
    data = extraction_data or {}

    def present(keys):
        for key in keys:
            value = str(data.get(key, "") or "").strip()
            if value and value.lower() not in ("unknown", "unknown device"):
                return True
        return False

    if present(manufacturer_keys) or present(model_keys):
        return []
    return ["Device manufacturer or model (not found in the extraction file)"]


def show_placeholder_preview(parent, rows, output_name="", allow_write=False):
    """Show PY_ values. allow_write is unused; preview is informational only."""
    from tkinter import Toplevel
    from tkinter import ttk
    import tkinter as tk

    result = {"ok": False}
    dialog = Toplevel(parent)
    dialog.title("Placeholder preview")
    dialog.transient(parent)
    dialog.grab_set()
    dialog.geometry("720x520")
    dialog.configure(bg=COLORS.get("bg", "#0b1e30"))

    ttk.Label(
        dialog,
        text="These PY_ values will be written into the template when you generate the report. Empty fields stay blank.",
        wraplength=680,
    ).pack(anchor="w", padx=14, pady=(12, 4))
    if output_name:
        ttk.Label(dialog, text=f"Suggested file name: {output_name}", style="Hint.TLabel").pack(
            anchor="w", padx=14, pady=(0, 6)
        )

    frame = ttk.Frame(dialog)
    frame.pack(fill="both", expand=True, padx=14, pady=6)
    text = tk.Text(frame, wrap="none", bg=COLORS.get("entry_bg", "#0a1724"),
                   fg=COLORS.get("entry_fg", "#e7f3fb"), insertbackground=COLORS.get("accent", "#2ec8e0"))
    yscroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
    text.configure(yscrollcommand=yscroll.set)
    text.pack(side="left", fill="both", expand=True)
    yscroll.pack(side="right", fill="y")

    text.insert("end", f"{'Placeholder':<18}  Value\n")
    text.insert("end", f"{'-'*18}  {'-'*48}\n")
    for key, value in rows:
        display = "" if value is None else str(value).replace("\n", " ").strip()
        if not display:
            display = "(empty)"
        if len(display) > 90:
            display = display[:87] + "..."
        text.insert("end", f"{key:<18}  {display}\n")
    text.configure(state="disabled")

    def accept():
        result["ok"] = True
        dialog.destroy()

    def cancel():
        result["ok"] = False
        dialog.destroy()

    buttons = ttk.Frame(dialog)
    buttons.pack(fill="x", padx=14, pady=10)
    ttk.Button(buttons, text="Close", style="Accent.TButton", command=cancel).pack(side="right")

    parent.wait_window(dialog)
    return result["ok"]


def mobile_preview_rows(data, officer_text="", image_date=""):
    return [
        ("PY_DFR", data.get("DFR_Num", "")),
        ("PY_CASENUMBER", data.get("Case_Number", "")),
        ("PY_EVIDENCE", data.get("evidence_ID", "")),
        ("PY_REQDATE", data.get("Request_Date", "")),
        ("PY_OWNER", data.get("Device_Owner", "")),
        ("PY_REQAGENCY", data.get("Request_Agency", "")),
        ("PY_REQOFF", officer_text),
        ("PY_EXAMINER", f"{data.get('Examiner_Title', '')} {data.get('Examiner_Name', '')}".strip()),
        ("PY_IMAGEDATE", image_date or data.get("formatted_date", "")),
        ("PY_MAN", data.get("device_manufacturer", "")),
        ("PY_MOD", data.get("device_model", "")),
        ("PY_COLOR", data.get("device_color", "")),
        ("PY_PHONE", data.get("Phone_Number", "")),
        ("PY_SERIAL", data.get("Serial_Number", "")),
        ("PY_IMEI", data.get("DEV_IMEI", "")),
        ("PY_CAPACITY", data.get("device_capacity", "")),
        ("PY_DEVNAME", data.get("Device_Name", "")),
        ("PY_ACCOUNT", data.get("Device_Account", "")),
        ("PY_ICCID", data.get("device_iccid", "")),
        ("PY_PASSCODE", data.get("device_passcode", "")),
        ("PY_CARRIER", data.get("device_carrier", "")),
        ("PY_OS", data.get("Device_OS", "")),
        ("PY_CBVER", data.get("cellebrite_version", "")),
        ("PY_GKVER", data.get("GrayKey_OS", "")),
    ]


def pc_preview_rows(data, officer_text="", image_date=""):
    return [
        ("PY_DFR", data.get("DFR_Num", "")),
        ("PY_CASENUMBER", data.get("Case_Number", "")),
        ("PY_EVIDENCE", data.get("evidence_ID", "")),
        ("PY_REQDATE", data.get("Request_Date", "")),
        ("PY_OWNER", data.get("Device_Owner", "")),
        ("PY_REQAGENCY", data.get("Request_Agency", "")),
        ("PY_REQOFF", officer_text),
        ("PY_EXAMINER", f"{data.get('Examiner_Title', '')} {data.get('Examiner_Name', '')}".strip()),
        ("PY_IMAGEDATE", image_date or data.get("formatted_date", "")),
        ("PY_DEVMAKE", data.get("device_PCMan", "")),
        ("PY_DEVMODEL", data.get("device_PCMod", "")),
        ("PY_PCSERIAL", data.get("device_PCSerial", "")),
        ("PY_COLOR", data.get("Device_Color", "")),
        ("PY_PASSCODE", data.get("device_password", "")),
        ("PY_HDMAKE", data.get("hd_make", "")),
        ("PY_HDMODEL", data.get("hd_model", "")),
        ("PY_HDSERIAL", data.get("hd_serial", "")),
        ("PY_CAPACITY", data.get("Device_Capacity", "")),
        ("PY_FTKVER", data.get("FTK_OS", "")),
        ("PY_TX1VER", data.get("TX1_OS", "")),
        ("PY_XWVER", data.get("xways_OS", "")),
        ("PY_DCVER", data.get("DC_OS", "")),
    ]


def apply_suggested_filename(app, module_label, model=""):
    widget = getattr(app, "output_filename", None)
    if widget is None:
        return ""
    try:
        current = widget.get().strip()
    except Exception:
        current = ""
    dfr = ""
    owner = ""
    try:
        if hasattr(app, "DFR_Num"):
            dfr = app.DFR_Num.get().strip()
        elif hasattr(app, "dfr_num"):
            dfr = app.dfr_num.get().strip()
    except Exception:
        pass
    try:
        owner = app.device_owner.get().strip()
    except Exception:
        pass
    return suggested_report_filename(dfr, module_label, owner, model)


def seed_save_location(app):
    widget = getattr(app, "save_location", None)
    if widget is None:
        return
    try:
        current = widget.get().strip()
    except Exception:
        current = ""
    if current:
        return
    try:
        widget.insert(0, default_export_dir())
    except Exception:
        pass
