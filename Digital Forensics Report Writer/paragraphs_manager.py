"""Load, save, and fill editable report paragraphs.

Factory wording lives in paragraphs_defaults.py (baked into the exe).
User edits are stored next to settings.json:
  frozen exe  -> %APPDATA%\\Digital Forensics Report Writer\\paragraphs.json
  from source -> paragraphs.json beside the scripts
"""
import json
import re
import shutil
import sys
from pathlib import Path

from paragraphs_defaults import DEFAULT_PARAGRAPHS
from ui_theme import APP_NAME, LEGACY_APP_NAME


REPORT_TYPES = [
    ("mobile_full", "Mobile — Full Exam"),
    ("mobile_portable", "Mobile — Portable Case"),
    ("pc_full", "PC — Full Exam"),
    ("pc_portable", "PC — Portable Case"),
    ("warrant", "Warrant Data Returns"),
]

# Shown in the editor list. Keys that are omitted fall back to the raw key name.
PARAGRAPH_META = {
    "mobile_full": {
        "one_a": ("Opening — Case Agent", "Used when the examiner role is Case Agent."),
        "one_b": ("Opening — Agency Assist", "Agency Assist and the device was not transferred."),
        "one_c": ("Opening — Agency Assist + transfer", "Agency Assist and custody was transferred."),
        "two_a": ("Airplane mode on", "Airplane mode is Yes."),
        "two_b": ("Airplane mode off", "Airplane mode is No."),
        "auth_c": ("Authority — Consent", "Legal authority is Consent."),
        "auth_sw": ("Authority — Search warrant", "Agency Assist and legal authority is Search Warrant."),
        "auth_i": ("Authority — Implied consent", "Legal authority is Implied Consent."),
        "auth_p": ("Authority — Parole", "Legal authority is Parole."),
        "three_a": ("Purpose — Agency Assist", "Role is Agency Assist."),
        "three_b": ("Purpose — Case Agent", "Role is Case Agent."),
        "four": ("Heading — Forensic Extraction", "Section heading (bold / underline)."),
        "five_cellebrite": ("Extraction — Cellebrite", "Single Cellebrite extraction."),
        "five_cellebrite_multiple": ("Extraction — Cellebrite, multiple", "More than one Cellebrite extraction."),
        "five_graykey": ("Extraction — GrayKey", "GrayKey extraction."),
        "six": ("Evidence storage", "Where the extraction is stored."),
        "seven": ("Heading — Forensic Processing", "Section heading (bold / underline)."),
        "eight_axiom": ("Processing — Axiom only", "Axiom is the only processing tool."),
        "eight_cellebrite": ("Processing — Cellebrite only", "Cellebrite is the only processing tool."),
        "eight_both": ("Processing — Axiom and Cellebrite", "Both tools, or all three tools."),
        "griffeye_para": ("Processing — Griffeye", "Added when Griffeye is checked."),
        "nine": ("Heading — Findings", "Section heading (bold / underline)."),
        "ten": ("Findings — media and files", "Used unless No Evidence Found is checked."),
        "eleventeen": ("Findings — apps and searches", "Used unless No Evidence Found is checked."),
        "twelveteen": ("Findings — AXIOM report intro", "Used unless No Evidence Found is checked."),
        "thirteen": ("Heading — AXIOM Digital Report", "Section heading (bold / underline)."),
        "fourteen": ("AXIOM — referral", "Used unless No Evidence Found is checked."),
        "fifteen": ("AXIOM — Case Information", "Used unless No Evidence Found is checked."),
        "sixteen": ("AXIOM — Artifacts intro", "Used unless No Evidence Found is checked."),
        "seventeen": ("Heading — Artifacts", "Section heading (bold / underline)."),
        "Paragraph_NoEv": ("No evidence found", "Used when No Evidence Found is checked."),
    },
    "mobile_portable": {
        "one_a": ("Opening — no transfer", "Device was not transferred."),
        "one_b": ("Opening — with transfer", "Custody was transferred."),
        "two": ("Device description", "Always included."),
        "auth_c": ("Authority — Consent", "Legal authority is Consent."),
        "auth_sw": ("Authority — Search warrant", "Legal authority is Search Warrant."),
        "auth_i": ("Authority — Implied consent", "Legal authority is Implied Consent."),
        "auth_p": ("Authority — Parole", "Legal authority is Parole."),
        "three_a": ("Airplane mode on", "Airplane mode is Yes."),
        "three_b": ("Airplane mode off", "Airplane mode is No."),
        "four": ("Purpose of examination", "Always included."),
        "five": ("Heading — Forensic Extraction", "Section heading (bold / underline)."),
        "six_cellebrite": ("Extraction — Cellebrite", "Single Cellebrite extraction."),
        "six_cellebrite_multiple": ("Extraction — Cellebrite, multiple", "More than one Cellebrite extraction."),
        "six_graykey": ("Extraction — GrayKey", "GrayKey extraction."),
        "seven_axiom": ("Processing — Axiom only", "Axiom is the only processing tool."),
        "seven_cellebrite": ("Processing — Cellebrite only", "Cellebrite is the only processing tool."),
        "seven_both": ("Processing — Axiom and Cellebrite", "Both tools selected."),
        "eight": ("Portable case delivered", "Always included."),
        "nine": ("No further examination", "Always included."),
        "ten": ("Evidence storage", "Always included."),
    },
    "pc_full": {
        "one_a_computer": ("Opening — Case Agent, computer", "Case Agent and device type is Computer."),
        "one_b_computer": ("Opening — Assist, computer", "Agency Assist, Computer, no transfer."),
        "one_c_computer": ("Opening — Assist + transfer, computer", "Agency Assist, Computer, transferred."),
        "one_a_loose": ("Opening — Case Agent, loose HDD", "Case Agent and device type is Loose Hard Drive."),
        "one_b_loose": ("Opening — Assist, loose HDD", "Agency Assist, Loose Hard Drive, no transfer."),
        "one_c_loose": ("Opening — Assist + transfer, loose HDD", "Agency Assist, Loose Hard Drive, transferred."),
        "one_a_storage": ("Opening — Case Agent, other storage", "Case Agent and device type is other storage."),
        "one_b_storage": ("Opening — Assist, other storage", "Agency Assist, other storage, no transfer."),
        "one_c_storage": ("Opening — Assist + transfer, other storage", "Agency Assist, other storage, transferred."),
        "auth_c": ("Authority — Consent", "Legal authority is Consent."),
        "auth_sw": ("Authority — Search warrant", "Agency Assist and legal authority is Search Warrant."),
        "auth_i": ("Authority — Implied consent", "Legal authority is Implied Consent."),
        "auth_p": ("Authority — Parole", "Legal authority is Parole."),
        "three_a": ("Purpose — Agency Assist", "Role is Agency Assist."),
        "three_b": ("Purpose — Case Agent", "Role is Case Agent."),
        "four": ("Heading — Forensic Extraction", "Section heading (bold / underline)."),
        "five_tx1": ("Imaging — TX1", "Acquisition tool is TX1."),
        "five_ftk": ("Imaging — FTK Imager", "Acquisition tool is FTK Imager."),
        "five_xways": ("Imaging — X-Ways", "Acquisition tool is X-Ways."),
        "five_dc": ("Imaging — Digital Collector", "Acquisition tool is Cellebrite Digital Collector."),
        "six": ("Evidence storage", "Always included."),
        "seven": ("Heading — Forensic Processing", "Section heading (bold / underline)."),
        "eight": ("Processing software", "Always included."),
        "griffeye_para": ("Processing — Griffeye", "Added when Griffeye is checked."),
        "nine": ("Heading — Findings", "Section heading (bold / underline)."),
        "ten": ("Findings — media and files", "Used unless No Evidence Found is checked."),
        "eleventeen": ("Findings — apps and searches", "Used unless No Evidence Found is checked."),
        "twelveteen": ("Findings — AXIOM report intro", "Used unless No Evidence Found is checked."),
        "thirteen": ("Heading — AXIOM Digital Report", "Section heading (bold / underline)."),
        "fourteen": ("AXIOM — referral", "Used unless No Evidence Found is checked."),
        "fifteen": ("AXIOM — Case Information", "Used unless No Evidence Found is checked."),
        "sixteen": ("AXIOM — Artifacts intro", "Used unless No Evidence Found is checked."),
        "seventeen": ("Heading — Artifacts", "Section heading (bold / underline)."),
        "Paragraph_NoEv": ("No evidence found", "Used when No Evidence Found is checked."),
    },
    "pc_portable": {
        "one_b_computer": ("Opening — computer, no transfer", "Computer and the device was not transferred."),
        "one_c_computer": ("Opening — computer + transfer", "Computer and custody was transferred."),
        "one_b_loose": ("Opening — loose HDD, no transfer", "Loose Hard Drive, not transferred."),
        "one_c_loose": ("Opening — loose HDD + transfer", "Loose Hard Drive, transferred."),
        "one_b_storage": ("Opening — other storage, no transfer", "Other storage device, not transferred."),
        "one_c_storage": ("Opening — other storage + transfer", "Other storage device, transferred."),
        "auth_c": ("Authority — Consent", "Legal authority is Consent."),
        "auth_sw": ("Authority — Search warrant", "Legal authority is Search Warrant."),
        "auth_i": ("Authority — Implied consent", "Legal authority is Implied Consent."),
        "auth_p": ("Authority — Parole", "Legal authority is Parole."),
        "three_a": ("Purpose of examination", "Always included."),
        "four": ("Heading — Forensic Extraction", "Section heading (bold / underline)."),
        "five_tx1": ("Imaging — TX1", "Acquisition tool is TX1."),
        "five_ftk": ("Imaging — FTK Imager", "Acquisition tool is FTK Imager."),
        "five_xways": ("Imaging — X-Ways", "Acquisition tool is X-Ways."),
        "five_dc": ("Imaging — Digital Collector", "Acquisition tool is Cellebrite Digital Collector."),
        "six": ("Evidence storage", "Always included."),
        "seven": ("Heading — Forensic Processing", "Section heading (bold / underline)."),
        "eight": ("Processing software", "Always included."),
        "nine": ("Portable case delivered", "Always included."),
        "ten": ("No further examination", "Always included."),
    },
    "warrant": {
        "intro_self": (
            "Warrant service and return",
            "Opening narrative for the warrant-return report. Field names on this report match the template tokens.",
        ),
    },
}

TOKEN_HELP = {
    "Request_Date": "Exam request date from the form",
    "Request_Title": "Requesting officer title",
    "Request_Officer": "Requesting officer full name",
    "Request_Officer_LastName": "Requesting officer last name only",
    "Request_Agency": "Requesting agency full name",
    "Request_Agency_Abbr": "Requesting agency abbreviation",
    "Request_Case": "Primary case offense / investigation type",
    "Device_Owner": "Device owner",
    "Examiner_Title": "Examiner title",
    "Examiner_Name": "Examiner name",
    "Examiner_Agency": "Examiner agency full name",
    "Examiner_Agency_Abbr": "Examiner agency abbreviation",
    "SW_Date": "Search warrant service date (Case Agent)",
    "Transfer_Date": "Date custody was transferred",
    "Transfer_Title": "Transferring officer title",
    "Transfer_Officer": "Transferring officer name",
    "Transfer_Agency": "Transferring officer agency",
    "article": "a / an chosen from the manufacturer",
    "device_manufacturer": "Mobile device manufacturer",
    "device_model": "Mobile device model",
    "device_PCMan": "Computer or drive manufacturer",
    "device_PCMod": "Computer or drive model",
    "device_capacity": "Drive or device capacity",
    "device_type": "Storage device type (other storage)",
    "hd_make": "Hard-drive manufacturer",
    "hd_model": "Hard-drive model",
    "source_device": "Narrative name for the exhibit (source device / source drive)",
    "formatted_date": "Acquisition or imaging date (and time when known)",
    "extraction_tool": "Acquisition tool name",
    "extraction_type": "Extraction or image type",
    "extraction_serial": "TX1 hardware serial number",
    "TX1_OS": "TX1 software version",
    "FTK_OS": "FTK Imager version",
    "xways_OS": "X-Ways version",
    "DC_OS": "Cellebrite Digital Collector version",
    "Forensic_Software": "Selected processing software name(s)",
    "PY_SERVEDATE": "Warrant service date",
    "PY_RETURNDATE": "Data return date",
    "PY_PROVIDER": "Service provider",
    "PY_ACCOUNTID": "Account identifier",
    "PY_DATASIZE": "Returned data size",
    "PY_LIMITSTART": "Time-frame start (if limited)",
    "PY_LIMITEND": "Time-frame end (if limited)",
    "PY_DFR": "DFR / WDR number",
    "PY_CASENUMBER": "Agency or lab case number",
    "PY_EXAMINER": "Examiner title and name",
    "PY_REQAGENCY": "Requesting agency",
    "PY_REQOFF": "Requesting officer title and name",
}

_TOKEN_RE = re.compile(r"\{([A-Za-z_][A-Za-z0-9_]*)\}")


class _SafeDict(dict):
    def __missing__(self, key):
        return "{" + key + "}"


def fill_paragraph(text, data):
    """Replace {Field} tokens. Unknown names are left as {Field} instead of crashing."""
    if not text:
        return ""
    try:
        return str(text).format_map(_SafeDict(data or {}))
    except (ValueError, IndexError):
        return str(text)


def paragraph_tokens(text):
    return _TOKEN_RE.findall(text or "")


def tokens_for_report(kind):
    seen = []
    for text in default_paragraphs(kind).values():
        for token in paragraph_tokens(text):
            if token not in seen:
                seen.append(token)
    extra = {
        "warrant": ["PY_LIMITSTART", "PY_LIMITEND", "PY_DFR", "PY_CASENUMBER", "PY_EXAMINER", "PY_REQAGENCY", "PY_REQOFF"],
    }
    for token in extra.get(kind, []):
        if token not in seen:
            seen.append(token)
    return seen


def validate_paragraph_text(text):
    """Return a list of problem strings, or an empty list if the braces look valid."""
    problems = []
    if text is None:
        return ["Paragraph text is missing."]
    stack = []
    for index, char in enumerate(text):
        if char == "{":
            stack.append(index)
        elif char == "}":
            if not stack:
                problems.append("A closing brace '}' has no matching '{'.")
                break
            start = stack.pop()
            name = text[start + 1:index]
            if not name:
                problems.append("Empty token {} is not allowed.")
            elif not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
                problems.append(f"Token {{{name}}} is not a valid field name.")
            if len(stack) > 1:
                problems.append("Tokens cannot be nested.")
                break
    else:
        if stack:
            problems.append("An opening brace '{' has no matching '}'.")
    return problems


def report_label(kind):
    for key, label in REPORT_TYPES:
        if key == kind:
            return label
    return kind


def paragraph_label(kind, key):
    meta = PARAGRAPH_META.get(kind, {}).get(key)
    return meta[0] if meta else key


def paragraph_when(kind, key):
    meta = PARAGRAPH_META.get(kind, {}).get(key)
    return meta[1] if meta else ""


def default_paragraphs(kind):
    return dict(DEFAULT_PARAGRAPHS.get(kind, {}))


def _appdata_dir():
    if sys.platform == "win32":
        base = os_appdata()
    else:
        from os import environ
        base = environ.get("XDG_CONFIG_HOME") or str(Path.home() / ".config")
    path = Path(base) / APP_NAME
    path.mkdir(parents=True, exist_ok=True)
    return path


def os_appdata():
    import os
    return os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")


def paragraphs_file():
    if getattr(sys, "frozen", False):
        path = _appdata_dir() / "paragraphs.json"
        if not path.is_file():
            legacy = Path(os_appdata()) / LEGACY_APP_NAME / "paragraphs.json"
            if legacy.is_file():
                try:
                    path.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(legacy, path)
                except Exception as exc:
                    print(f"Could not migrate paragraphs.json: {exc}")
        return path
    return Path(__file__).resolve().parent / "paragraphs.json"


def _read_overrides():
    path = paragraphs_file()
    try:
        if path.is_file():
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
    except Exception as exc:
        print(f"Error reading paragraphs {path}: {exc}")
    return {}


def _write_overrides(data):
    path = paragraphs_file()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        return True
    except Exception as exc:
        print(f"Error saving paragraphs {path}: {exc}")
        return False


def load_paragraphs(kind):
    merged = default_paragraphs(kind)
    overrides = _read_overrides().get(kind) or {}
    if isinstance(overrides, dict):
        for key, value in overrides.items():
            if key in merged and isinstance(value, str):
                merged[key] = value
    return merged


def is_modified(kind, key):
    current = load_paragraphs(kind).get(key, "")
    factory = default_paragraphs(kind).get(key, "")
    return current != factory


def save_paragraph(kind, key, text):
    factory = default_paragraphs(kind)
    if key not in factory:
        return False
    data = _read_overrides()
    bucket = data.setdefault(kind, {})
    if text == factory[key]:
        bucket.pop(key, None)
        if not bucket:
            data.pop(kind, None)
    else:
        bucket[key] = text
    return _write_overrides(data)


def revert_paragraph(kind, key):
    data = _read_overrides()
    bucket = data.get(kind) or {}
    if key in bucket:
        bucket.pop(key, None)
        if not bucket:
            data.pop(kind, None)
        return _write_overrides(data)
    return True


def revert_report(kind):
    data = _read_overrides()
    if kind in data:
        data.pop(kind, None)
        return _write_overrides(data)
    return True


def revert_all():
    path = paragraphs_file()
    try:
        if path.is_file():
            path.unlink()
        return True
    except Exception as exc:
        print(f"Error removing paragraphs {path}: {exc}")
        return False


def refresh_open_report_windows(root):
    """Reload paragraph dictionaries on any open report window."""
    if root is None:
        return
    seen = set()
    stack = [root]
    while stack:
        widget = stack.pop()
        ident = id(widget)
        if ident in seen:
            continue
        seen.add(ident)
        kind = getattr(widget, "_paragraph_kind", None)
        if kind:
            try:
                widget.paragraphs = load_paragraphs(kind)
            except Exception:
                pass
        try:
            stack.extend(widget.winfo_children())
        except Exception:
            pass


# Keep a bundled copy next to the defaults so a source checkout can seed later.
def bundled_defaults_exist():
    return Path(__file__).resolve().parent.joinpath("paragraphs_defaults.py").is_file()
