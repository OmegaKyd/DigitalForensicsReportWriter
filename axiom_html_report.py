"""Parse Magnet AXIOM HTML or XML tagged-item exports and pre-check artifact boxes.

Folder names and the frameset/XML filename do not have to match Magnet defaults.
HTML path reads Resources/nav.html. XML path reads Export.xml (iefExport).
Neither path opens attachments or media files.
"""
from __future__ import annotations

import html
import json
import os
import re
from html.parser import HTMLParser
from pathlib import Path
from xml.etree import ElementTree as ET

MEDIA_PARENTS = ("Pictures", "Videos")
MEDIA_SUBTAGS = ("Child Pornography", "Child Erotica", "Age Difficult")

SKIP_DIR_NAMES = {
    "attachments",
    "chat preview report",
    "resources",
}

CP_PHRASES = (
    "child pornography",
    "child porn",
    "child sexual abuse material",
    "child abuse material",
    "csam",
)
CE_PHRASES = (
    "child erotica",
    "child exploitive",
    "child exploitative",
)
AD_PHRASES = (
    "age difficult",
    "age questionable",
)
CP_TOKENS = {"cp", "csam", "cam"}
CE_TOKENS = {"ce"}
AD_TOKENS = {"ad"}

DROP_HINT = "Drag & Drop AXIOM HTML or XML report (optional) — or click to browse"
VERSION_RE = re.compile(r"(\d+\.\d+\.\d+(?:\.\d+)?)")


class _AnchorCollector(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.events = []
        self._in_a = False
        self._href = ""
        self._parts = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "center" and (attrs.get("id") or "").lower() == "artifactgroupname":
            self.events.append(("group_start", ""))
        elif tag == "a":
            self._in_a = True
            self._href = attrs.get("href") or ""
            self._parts = []

    def handle_endtag(self, tag):
        if tag == "a" and self._in_a:
            text = html.unescape("".join(self._parts)).strip()
            self.events.append(("link", text, self._href))
            self._in_a = False
            self._href = ""
            self._parts = []

    def handle_data(self, data):
        text = data.strip()
        if not text:
            return
        if self._in_a:
            self._parts.append(data)
        elif self.events and self.events[-1][0] == "group_start" and self.events[-1][1] == "":
            self.events[-1] = ("group_start", text)


def _read_text(path):
    with open(path, "r", encoding="utf-8-sig", errors="ignore") as handle:
        return handle.read()


def _is_html(path):
    return Path(path).suffix.lower() in {".html", ".htm"}


def _looks_like_nav(text):
    lowered = text.lower()
    return "artifactgroupname" in lowered or "magnet axiom report" in lowered


def _looks_like_frameset(text):
    lowered = text.lower()
    return "<frameset" in lowered and "nav.html" in lowered


def _nav_from_frameset(report_path, text):
    match = re.search(r'src=["\']([^"\']*nav\.html)["\']', text, re.I)
    if not match:
        return None
    src = match.group(1).replace("\\", "/").lstrip("./")
    candidate = Path(report_path).parent / src
    if candidate.is_file():
        return candidate
    return None


def _iter_html_limited(root, max_files=80):
    root = Path(root)
    stack = [(root, 0)]
    seen = 0
    while stack:
        current, depth = stack.pop()
        if depth > 4:
            continue
        try:
            entries = list(current.iterdir())
        except OSError:
            continue
        dirs = []
        for entry in entries:
            name = entry.name.lower()
            if entry.is_dir():
                if name in SKIP_DIR_NAMES and depth > 0:
                    continue
                dirs.append(entry)
            elif _is_html(entry):
                yield entry
                seen += 1
                if seen >= max_files:
                    return
        for directory in reversed(dirs):
            stack.append((directory, depth + 1))


def _is_xml(path):
    return Path(path).suffix.lower() == ".xml"


def _looks_like_axiom_xml(text):
    lowered = (text or "")[:2000].lower()
    return "<iefexport" in lowered


def _normalize_version(value):
    text = html.unescape(str(value or "")).strip()
    if not text:
        return ""
    match = VERSION_RE.search(text)
    return match.group(1) if match else text


def _version_from_summary_json(path):
    try:
        payload = json.loads(_read_text(path))
    except Exception:
        return ""
    for item in payload.get("summary") or []:
        if str(item.get("name") or "").strip().lower() == "axiom version":
            return _normalize_version(item.get("value"))
    return ""


def find_export_summary(start):
    start = Path(start)
    roots = []
    if start.is_file():
        roots.extend([start.parent, start.parent.parent])
    else:
        roots.append(start)
        roots.append(start / "Export")
    seen = set()
    for root in roots:
        if root in seen or not root.exists():
            continue
        seen.add(root)
        candidate = root / "ExportSummary.json"
        if candidate.is_file():
            return candidate
        try:
            for child in root.iterdir():
                if child.is_file() and child.name.lower() == "exportsummary.json":
                    return child
        except OSError:
            continue
    return None


def extract_axiom_version(*paths):
    for path in paths:
        if not path:
            continue
        target = Path(path)
        summary = find_export_summary(target)
        if summary:
            version = _version_from_summary_json(summary)
            if version:
                return version
        if target.is_file() and _is_xml(target):
            try:
                text = _read_text(target)[:4000]
            except OSError:
                text = ""
            match = re.search(r'iefVersion="([^"]+)"', text, re.I)
            if match:
                version = _normalize_version(match.group(1))
                if version:
                    return version
        if target.is_file() and _is_html(target):
            try:
                text = _read_text(target)[:4000]
            except OSError:
                text = ""
            match = re.search(r"Magnet AXIOM Report\s*v?([\d.]+)", text, re.I)
            if match:
                return _normalize_version(match.group(1))
    return ""


def find_axiom_xml(path):
    if not path:
        return None
    target = Path(path)
    if not target.exists():
        return None
    if target.is_file():
        if target.suffix.lower() == ".zip":
            return None
        if _is_xml(target):
            try:
                text = _read_text(target)
            except OSError:
                return None
            return target if _looks_like_axiom_xml(text) else None
        summary = find_export_summary(target)
        if summary:
            sibling = summary.parent / "Export.xml"
            if sibling.is_file():
                return sibling
        return find_axiom_xml(target.parent)

    preferred = []
    others = []
    try:
        entries = list(target.iterdir())
    except OSError:
        entries = []
    for child in entries:
        if child.is_file() and _is_xml(child):
            if child.name.lower() == "export.xml":
                preferred.append(child)
            else:
                others.append(child)
    for candidate in preferred + others:
        try:
            if _looks_like_axiom_xml(_read_text(candidate)):
                return candidate
        except OSError:
            continue
    nested = target / "Export"
    if nested.is_dir():
        found = find_axiom_xml(nested)
        if found:
            return found
    for child in entries:
        if child.is_dir() and child.name.lower() not in SKIP_DIR_NAMES:
            found = find_axiom_xml(child)
            if found:
                return found
    return None


def find_axiom_source(path):
    xml_path = find_axiom_xml(path)
    if xml_path:
        return "xml", xml_path
    html_path = find_axiom_report(path)
    if html_path:
        return "html", html_path
    return None, None


def find_axiom_report(path):
    """Return a usable HTML file (frameset or nav) or None."""
    if not path:
        return None
    target = Path(path)
    if not target.exists():
        return None

    if target.is_file():
        if target.suffix.lower() == ".zip":
            return None
        if _is_html(target):
            text = _read_text(target)
            if _looks_like_nav(text) or _looks_like_frameset(text):
                return target
            sibling_nav = target.parent / "Resources" / "nav.html"
            if sibling_nav.is_file():
                return target
            parent_nav = target.parent / "nav.html"
            if parent_nav.is_file():
                return parent_nav
            return target
        return find_axiom_report(target.parent)

    preferred_names = {"report.html", "report.htm", "nav.html"}
    for child in target.iterdir() if target.is_dir() else []:
        if child.is_file() and child.name.lower() in preferred_names:
            return child
    nested = target / "Export"
    if nested.is_dir():
        found = find_axiom_report(nested)
        if found:
            return found
    for html_path in _iter_html_limited(target):
        if html_path.name.lower() in preferred_names:
            return html_path
    for html_path in _iter_html_limited(target):
        try:
            text = _read_text(html_path)
        except OSError:
            continue
        if _looks_like_frameset(text) or _looks_like_nav(text):
            return html_path
    return None


def _resolve_nav(report_path):
    path = Path(report_path)
    text = _read_text(path)
    if _looks_like_nav(text):
        return path, text
    if _looks_like_frameset(text):
        nav = _nav_from_frameset(path, text)
        if nav and nav.is_file():
            return nav, _read_text(nav)
    for candidate in (
        path.parent / "Resources" / "nav.html",
        path.parent / "nav.html",
        path.parent / "resources" / "nav.html",
    ):
        if candidate.is_file():
            return candidate, _read_text(candidate)
    return path, text


def parse_nav_lists(nav_text):
    parser = _AnchorCollector()
    parser.feed(nav_text)
    artifacts = []
    tags = []
    current_group = None
    in_tags = False
    seen_artifacts = set()
    seen_tags = set()

    for event in parser.events:
        kind = event[0]
        if kind == "group_start":
            current_group = event[1].strip()
            in_tags = current_group.lower() == "tags"
        elif kind == "link":
            label = event[1].strip()
            if not label:
                continue
            lowered = label.lower()
            if lowered in {"case information", "title page", "case overview", "evidence overview", "media categorization summary"}:
                continue
            if in_tags or (event[2] or "").lower().endswith("_tag.html"):
                if label not in seen_tags:
                    tags.append(label)
                    seen_tags.add(label)
            elif current_group:
                if label not in seen_artifacts:
                    artifacts.append(label)
                    seen_artifacts.add(label)
    return artifacts, tags


def _normalize_name(name):
    text = html.unescape(name or "").strip()
    text = text.replace("_", "/")
    text = re.sub(r"\s+", " ", text)
    return text.lower()


def _name_keys(name):
    raw = html.unescape(name or "").strip()
    keys = {raw.lower(), _normalize_name(raw)}
    keys.add(raw.replace("_", "/").lower())
    keys.add(raw.replace("/", "_").lower())
    keys.add(re.sub(r"\s+", " ", raw).lower())
    return {key for key in keys if key}


def _organize_selected(selected_artifacts):
    selected = list(selected_artifacts or [])
    organized = []
    for name in selected:
        if " -- " in name or name in MEDIA_PARENTS:
            continue
        organized.append(name)
    for parent in MEDIA_PARENTS:
        if parent in selected:
            organized.append(parent)
            for name in selected:
                if name.startswith(f"{parent} -- "):
                    organized.append(name)
    for name in selected:
        if name not in organized:
            organized.append(name)
    return organized


def _platforms_for_device_class(device_class):
    key = (device_class or "mobile").strip().lower()
    if key in ("pc", "computer/storage", "computer", "storage"):
        return ("Computer", "macOS", "Chromebook", "Cloud", "Refined Results")
    if key in ("warrant", "sw", "search warrant", "warrant return"):
        return ("Cloud", "Android", "iOS", "Computer", "macOS", "Chromebook", "Refined Results")
    if key == "cloud":
        return ("Cloud", "Refined Results")
    return ("Android", "iOS", "Cloud", "Refined Results")


def _preferred_platform(device_class, device_type=None):
    platforms = _platforms_for_device_class(device_class)
    hint = str(device_type or "").strip()
    for name in platforms:
        if name.lower() == hint.lower():
            return name
    lowered = hint.lower()
    key = (device_class or "mobile").strip().lower()
    if lowered in ("ios", "iphone", "ipad", "ipod") and "iOS" in platforms:
        return "iOS"
    if lowered in ("macos", "mac", "os x", "osx") and "macOS" in platforms:
        return "macOS"
    if key in ("warrant", "sw", "search warrant", "warrant return", "cloud") and "Cloud" in platforms:
        return "Cloud"
    if "Computer" in platforms:
        return "Computer"
    return platforms[0]


def _load_catalog_platforms():
    import json
    import sys

    candidates = [Path(__file__).resolve().parent / "magnet_artifacts.json"]
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        candidates.insert(0, Path(sys._MEIPASS) / "magnet_artifacts.json")
    path = next((item for item in candidates if item.is_file()), None)
    if path is None:
        raise FileNotFoundError("magnet_artifacts.json was not found")
    with open(path, "r", encoding="utf-8") as handle:
        payload = json.load(handle)
    platforms = payload.get("platforms") if isinstance(payload, dict) else payload
    if not isinstance(platforms, dict):
        raise ValueError("magnet_artifacts.json is missing a platforms object")
    return platforms


def build_catalog_lookup(platforms):
    catalog = _load_catalog_platforms()
    lookup = {}
    for platform in platforms:
        for item in catalog.get(platform) or ():
            name = item.get("name") or ""
            if not name:
                continue
            record = {"name": name, "platform": platform, "category": item.get("category") or ""}
            for key in _name_keys(name):
                lookup.setdefault(key, []).append(record)
    return lookup


def match_artifact_name(name, lookup, preferred_platforms):
    candidates = []
    for key in _name_keys(name):
        candidates.extend(lookup.get(key) or [])
    if not candidates:
        return None
    preferred = list(preferred_platforms or [])
    for platform in preferred:
        for record in candidates:
            if record["platform"] == platform:
                return record
    return candidates[0]


def _tokens(text):
    return set(re.findall(r"[a-z0-9]+", (text or "").lower()))


def _plain_cell(html_fragment):
    text = re.sub(r"<li[^>]*>", "\n", html_fragment or "", flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _tag_page_title(text, path=None):
    match = re.search(r'id="page-title"[^>]*>([^<]+)', text or "", re.I)
    if match:
        return html.unescape(match.group(1)).strip()
    match = re.search(r"<title>([^<]+)</title>", text or "", re.I)
    if match:
        return html.unescape(match.group(1)).strip()
    if path:
        return Path(path).stem.replace("_tag", "").replace("_", "/")
    return ""


def _tbody_html(text):
    match = re.search(r"<tbody\b[^>]*>(.*?)</tbody>", text or "", re.I | re.S)
    return match.group(1) if match else (text or "")


def _count_hit_rows(text):
    return len(re.findall(r"<tr\b", _tbody_html(text), re.I))


def _lookup_count(counts, name):
    if not counts or not name:
        return None
    if name in counts:
        return counts[name]
    needle = _normalize_name(name)
    for key, value in counts.items():
        if _normalize_name(key) == needle:
            return value
    return None


def parse_tag_page(path):
    text = _read_text(path)
    tag_name = _tag_page_title(text, path)
    artifact_types = []
    seen_types = set()
    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", _tbody_html(text), re.I | re.S)
    for row in rows:
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.I | re.S)
        if len(cells) < 5:
            continue
        artifact_type = _plain_cell(cells[4])
        if artifact_type and artifact_type.lower() not in seen_types:
            artifact_types.append(artifact_type)
            seen_types.add(artifact_type.lower())
    return {
        "tag": tag_name,
        "artifact_types": artifact_types,
        "count": len(rows),
        "path": str(path),
    }


def collect_html_counts(nav_path, lookup=None, preferred=None):
    artifact_counts = {}
    tag_counts = {}
    webpages = _webpages_dir(nav_path)
    if not webpages.is_dir():
        return artifact_counts, tag_counts
    pages = list(webpages.glob("*.html")) + list(webpages.glob("*.htm"))
    for page in pages:
        lowered = page.name.lower()
        if lowered.endswith("_tag.html") or lowered.endswith("_tag.htm"):
            info = parse_tag_page(page)
            label = info.get("tag") or ""
            if label:
                tag_counts[label] = int(info.get("count") or 0)
            continue
        try:
            text = _read_text(page)
        except OSError:
            continue
        title = _tag_page_title(text, page)
        if not title:
            continue
        if title.lower() in {"case overview", "evidence overview", "media categorization summary", "title page", "case information"}:
            continue
        key = title
        if lookup is not None:
            record = match_artifact_name(title, lookup, preferred)
            if record:
                key = record["name"]
        count = _count_hit_rows(text)
        artifact_counts[key] = artifact_counts.get(key, 0) + count
    return artifact_counts, tag_counts


def _webpages_dir(nav_path):
    parent = Path(nav_path).parent
    for candidate in (parent / "webpages", parent / "Webpages", parent):
        if candidate.is_dir() and any(candidate.glob("*_tag.html")):
            return candidate
    return parent / "webpages"


def collect_examiner_tags(nav_path, raw_tags, lookup, preferred):
    """Map catalog artifact / media-subtag keys to examiner tag labels from the HTML export."""
    artifact_tags = {}
    subtag_tags = {}

    def add_tag(bucket, key, label):
        if not key or not label:
            return
        group = bucket.setdefault(key, [])
        needle = _normalize_name(label)
        for index, existing in enumerate(group):
            if _normalize_name(existing) == needle:
                if "/" in label and "_" in existing:
                    group[index] = label
                return
        group.append(label)

    pages = []
    webpages = _webpages_dir(nav_path)
    if webpages.is_dir():
        pages = sorted(webpages.glob("*_tag.html")) + sorted(webpages.glob("*_tag.htm"))
    parsed_labels = set()
    for page in pages:
        info = parse_tag_page(page)
        label = info.get("tag") or ""
        if label:
            parsed_labels.add(label.lower())
        mapped = classify_media_tag(label)
        if mapped:
            for item in mapped:
                add_tag(subtag_tags, item, label)
            continue
        types = info.get("artifact_types") or []
        if not types and label:
            types = [label]
        attached = False
        for artifact_type in types:
            record = match_artifact_name(artifact_type, lookup, preferred)
            if record:
                add_tag(artifact_tags, record["name"], label)
                attached = True
        if not attached and label:
            record = match_artifact_name(label, lookup, preferred)
            if record:
                add_tag(artifact_tags, record["name"], label)

    for tag in raw_tags or []:
        if tag.lower() in parsed_labels:
            continue
        mapped = classify_media_tag(tag)
        if mapped:
            for item in mapped:
                add_tag(subtag_tags, item, tag)
            continue
        record = match_artifact_name(tag, lookup, preferred)
        if record:
            add_tag(artifact_tags, record["name"], tag)

    return artifact_tags, subtag_tags


def classify_media_tag(tag_name):
    raw = html.unescape(tag_name or "").strip()
    if not raw:
        return None
    lowered = raw.lower()
    tokens = _tokens(lowered)

    parent = None
    if "pictures" in tokens or "picture" in tokens:
        parent = "Pictures"
    elif "videos" in tokens or "video" in tokens:
        parent = "Videos"

    kind = None
    if any(phrase in lowered for phrase in CP_PHRASES) or (tokens & CP_TOKENS):
        kind = "Child Pornography"
    elif any(phrase in lowered for phrase in CE_PHRASES) or (tokens & CE_TOKENS):
        kind = "Child Erotica"
    elif any(phrase in lowered for phrase in AD_PHRASES) or (tokens & AD_TOKENS):
        kind = "Age Difficult"

    if not kind:
        return None
    if parent:
        return [f"{parent} -- {kind}"]
    return [f"{media} -- {kind}" for media in MEDIA_PARENTS]


def _split_tag_labels(text):
    raw = html.unescape(text or "").strip()
    if not raw:
        return []
    parts = [part.strip() for part in re.split(r"\s*,\s*", raw) if part.strip()]
    return parts or [raw]


def _catalog_context(device_class, device_type=None):
    platforms = _platforms_for_device_class(device_class)
    preferred = [_preferred_platform(device_class, device_type)]
    for name in platforms:
        if name not in preferred:
            preferred.append(name)
    lookup = build_catalog_lookup(preferred)
    return preferred, lookup


def _match_hits(raw_artifacts, raw_tags, artifact_tags, subtag_tags, preferred, lookup):
    matched = []
    sources = {}
    unmatched_artifacts = []
    seen = set()
    for label in raw_artifacts:
        record = match_artifact_name(label, lookup, preferred)
        if not record:
            unmatched_artifacts.append(label)
            continue
        name = record["name"]
        if name in seen:
            continue
        seen.add(name)
        matched.append(name)
        sources[name] = record["platform"]

    subtags = []
    unmatched_tags = []
    seen_sub = set()
    for item, labels in subtag_tags.items():
        parent, kind = item.split(" -- ", 1)
        if kind not in MEDIA_SUBTAGS:
            continue
        if item not in seen_sub:
            subtags.append(item)
            seen_sub.add(item)
        if parent in MEDIA_PARENTS and parent not in seen:
            matched.append(parent)
            seen.add(parent)
            if parent not in sources:
                sources[parent] = preferred[0]
    classified = set()
    for labels in subtag_tags.values():
        classified.update(name.lower() for name in labels)
    for tag in raw_tags:
        if tag.lower() in classified:
            continue
        if match_artifact_name(tag, lookup, preferred) is not None:
            continue
        attached = any(tag in names for names in artifact_tags.values())
        if not attached:
            unmatched_tags.append(tag)

    report_tags = dict(artifact_tags)
    report_tags.update(subtag_tags)
    return {
        "matched": _organize_selected(matched),
        "subtags": subtags,
        "sources": sources,
        "report_tags": report_tags,
        "unmatched_artifacts": unmatched_artifacts,
        "unmatched_tags": unmatched_tags,
    }


def parse_axiom_xml_report(path, device_class="mobile", device_type=None):
    found = find_axiom_xml(path)
    if not found:
        raise ValueError(
            "Could not find an AXIOM XML export in that location. "
            "Select Export.xml (any name) or the export folder that contains it."
        )
    try:
        tree = ET.parse(found)
        root = tree.getroot()
    except Exception as exc:
        raise ValueError(f"Could not read the AXIOM XML export:\n{exc}") from exc
    if (root.tag or "").split("}")[-1].lower() != "iefexport":
        raise ValueError("The selected file is not a Magnet AXIOM XML export (missing iefExport).")

    artifacts_node = None
    for child in root:
        if (child.tag or "").split("}")[-1].lower() == "artifacts":
            artifacts_node = child
            break
    if artifacts_node is None:
        raise ValueError("The AXIOM XML export does not contain an Artifacts list.")

    preferred, lookup = _catalog_context(device_class, device_type)
    raw_artifacts = []
    raw_tags = []
    seen_artifacts = set()
    seen_tags = set()
    artifact_tags = {}
    subtag_tags = {}
    artifact_counts = {}
    tag_counts = {}

    def add_label(bucket, key, label):
        if not key or not label:
            return
        group = bucket.setdefault(key, [])
        needle = _normalize_name(label)
        for index, existing in enumerate(group):
            if _normalize_name(existing) == needle:
                if "/" in label and "_" in existing:
                    group[index] = label
                return
        group.append(label)

    for artifact in artifacts_node:
        if (artifact.tag or "").split("}")[-1].lower() != "artifact":
            continue
        artifact_name = html.unescape(artifact.get("name") or "").strip()
        if not artifact_name:
            continue
        if artifact_name not in seen_artifacts:
            raw_artifacts.append(artifact_name)
            seen_artifacts.add(artifact_name)
        record = match_artifact_name(artifact_name, lookup, preferred)
        catalog_name = record["name"] if record else None
        hits = None
        for child in artifact:
            if (child.tag or "").split("}")[-1].lower() == "hits":
                hits = child
                break
        if hits is None:
            continue
        seen_items = set()
        for hit in hits:
            if (hit.tag or "").split("}")[-1].lower() != "hit":
                continue
            item_id = hit.get("sequenceNumber") or ""
            hit_labels = []
            for fragment in hit:
                if (fragment.tag or "").split("}")[-1].lower() != "fragment":
                    continue
                fname = (fragment.get("name") or "").strip().lower()
                if fname == "item id" and (fragment.text or "").strip():
                    item_id = (fragment.text or "").strip()
                elif fname == "tags":
                    hit_labels.extend(_split_tag_labels(fragment.text))
            unique_key = item_id or f"hit-{len(seen_items)}"
            if unique_key in seen_items:
                continue
            seen_items.add(unique_key)
            for label in hit_labels:
                if label not in seen_tags:
                    raw_tags.append(label)
                    seen_tags.add(label)
                tag_counts[label] = tag_counts.get(label, 0) + 1
                mapped = classify_media_tag(label)
                if mapped:
                    for item in mapped:
                        add_label(subtag_tags, item, label)
                    continue
                if catalog_name:
                    add_label(artifact_tags, catalog_name, label)
                else:
                    fallback = match_artifact_name(label, lookup, preferred)
                    if fallback:
                        add_label(artifact_tags, fallback["name"], label)
        count_key = catalog_name or artifact_name
        artifact_counts[count_key] = artifact_counts.get(count_key, 0) + len(seen_items)

    parsed = _match_hits(raw_artifacts, raw_tags, artifact_tags, subtag_tags, preferred, lookup)
    parsed.update({
        "report_path": str(found),
        "nav_path": str(found),
        "source_kind": "xml",
        "raw_artifacts": raw_artifacts,
        "raw_tags": raw_tags,
        "artifact_counts": artifact_counts,
        "tag_counts": tag_counts,
        "axiom_version": extract_axiom_version(found, path),
    })
    return parsed


def parse_axiom_report(path, device_class="mobile", device_type=None):
    kind, found = find_axiom_source(path)
    if kind == "xml":
        return parse_axiom_xml_report(found, device_class=device_class, device_type=device_type)
    if kind == "html":
        return parse_axiom_html_report(found, device_class=device_class, device_type=device_type)
    raise ValueError(
        "Could not find an AXIOM HTML or XML report in that location. "
        "Select Report.html, Export.xml, or the extracted export folder."
    )


def parse_axiom_html_report(path, device_class="mobile", device_type=None):
    found = find_axiom_report(path)
    if not found:
        raise ValueError(
            "Could not find an AXIOM HTML report in that location. "
            "Select Report.html (any name) or the export folder that contains it."
        )
    nav_path, nav_text = _resolve_nav(found)
    if not _looks_like_nav(nav_text):
        raise ValueError(
            "The selected file does not look like a Magnet AXIOM HTML report "
            "(missing the artifact navigation list)."
        )

    raw_artifacts, raw_tags = parse_nav_lists(nav_text)
    platforms = _platforms_for_device_class(device_class)
    preferred = [_preferred_platform(device_class, device_type)]
    for name in platforms:
        if name not in preferred:
            preferred.append(name)
    lookup = build_catalog_lookup(preferred)

    matched = []
    sources = {}
    unmatched_artifacts = []
    seen = set()
    for label in raw_artifacts:
        record = match_artifact_name(label, lookup, preferred)
        if not record:
            unmatched_artifacts.append(label)
            continue
        name = record["name"]
        if name in seen:
            continue
        seen.add(name)
        matched.append(name)
        sources[name] = record["platform"]

    artifact_tags, subtag_tags = collect_examiner_tags(nav_path, raw_tags, lookup, preferred)
    subtags = []
    unmatched_tags = []
    seen_sub = set()
    for item, labels in subtag_tags.items():
        parent, label = item.split(" -- ", 1)
        if label not in MEDIA_SUBTAGS:
            continue
        if item not in seen_sub:
            subtags.append(item)
            seen_sub.add(item)
        if parent in MEDIA_PARENTS and parent not in seen:
            matched.append(parent)
            seen.add(parent)
            if parent not in sources:
                sources[parent] = preferred[0]
    classified = set()
    for labels in subtag_tags.values():
        classified.update(name.lower() for name in labels)
    for tag in raw_tags:
        if tag.lower() in classified:
            continue
        if match_artifact_name(tag, lookup, preferred) is not None:
            continue
        attached = any(tag in names for names in artifact_tags.values())
        if not attached:
            unmatched_tags.append(tag)

    report_tags = dict(artifact_tags)
    report_tags.update(subtag_tags)
    artifact_counts, tag_counts = collect_html_counts(nav_path, lookup=lookup, preferred=preferred)

    return {
        "report_path": str(found),
        "nav_path": str(nav_path),
        "source_kind": "html",
        "raw_artifacts": raw_artifacts,
        "raw_tags": raw_tags,
        "matched": _organize_selected(matched),
        "subtags": subtags,
        "sources": sources,
        "report_tags": report_tags,
        "artifact_counts": artifact_counts,
        "tag_counts": tag_counts,
        "unmatched_artifacts": unmatched_artifacts,
        "unmatched_tags": unmatched_tags,
        "axiom_version": extract_axiom_version(nav_path, found, path),
    }


def _selection_from_parse(result):
    names = list(result.get("matched") or [])
    sources = dict(result.get("sources") or {})
    for sub in result.get("subtags") or []:
        parent = sub.split(" -- ", 1)[0]
        if parent in MEDIA_PARENTS and parent not in names:
            names.append(parent)
        if sub not in names:
            names.append(sub)
    return _organize_selected(names), sources


def apply_parse_to_form(parent, result):
    incoming, incoming_sources = _selection_from_parse(result)
    parent.selected_artifacts = _organize_selected(incoming)
    parent.selected_artifact_sources = dict(incoming_sources)
    parent.axiom_report_path = result.get("report_path") or ""
    parent.axiom_report_parse = result
    parent.axiom_version = result.get("axiom_version") or ""
    parent.axiom_report_tags = dict(result.get("report_tags") or {})
    tag_counts = {}
    for key, value in (result.get("tag_counts") or {}).items():
        try:
            tag_counts[key] = int(value)
        except (TypeError, ValueError):
            continue
    parent.axiom_tag_counts = tag_counts
    if hasattr(parent, "update_artifacts_count_label"):
        try:
            parent.update_artifacts_count_label()
        except Exception:
            pass


def summarize_parse(result):
    lines = []
    matched = result.get("matched") or []
    subtags = result.get("subtags") or []
    found = len(result.get("raw_artifacts") or [])
    lines.append(f"{len(matched)} of {found} artifact types matched the catalog.")
    if matched:
        lines.append("Matched artifacts:")
        lines.extend(f"  • {name}" for name in matched)
    if subtags:
        lines.append("Media tags:")
        lines.extend(f"  • {item.replace(' -- ', ' — ')}" for item in subtags)
    else:
        lines.append("No Child Pornography / Child Erotica / Age Difficult tags were mapped.")
    unmatched_a = result.get("unmatched_artifacts") or []
    unmatched_t = result.get("unmatched_tags") or []
    if unmatched_a:
        lines.append("Unmatched artifact names:")
        lines.extend(f"  • {name}" for name in unmatched_a)
    if unmatched_t:
        lines.append("Other custom tags (not mapped to CP/CE/AD):")
        lines.extend(f"  • {name}" for name in unmatched_t)
    return "\n".join(lines)


def _fill_listbox(listbox, items, empty_text):
    listbox.delete(0, "end")
    if items:
        for item in items:
            listbox.insert("end", item)
    else:
        listbox.insert("end", empty_text)


def _imported_tags_for(name, result):
    report_tags = result.get("report_tags") or {}
    labels = list(report_tags.get(name) or [])
    if name in MEDIA_PARENTS:
        for item in result.get("subtags") or []:
            if not item.startswith(f"{name} -- "):
                continue
            for label in report_tags.get(item) or ():
                if label not in labels:
                    labels.append(label)
    return labels


def _count_suffix(count, noun="item"):
    if count is None:
        return ""
    try:
        number = int(count)
    except (TypeError, ValueError):
        return ""
    label = noun if number == 1 else f"{noun}s"
    return f"  ({number} {label})"


def show_parse_dialog(parent, result):
    """Readable confirmation of parsed AXIOM artifacts and media tags."""
    import tkinter as tk
    from tkinter import ttk
    from ui_theme import COLORS, apply_theme

    matched = list(result.get("matched") or [])
    report_tags = result.get("report_tags") or {}
    artifact_counts = result.get("artifact_counts") or {}
    tag_counts = result.get("tag_counts") or {}
    art_blocks = [(name, _imported_tags_for(name, result)) for name in matched]
    imported_count = sum(len(labels) for _name, labels in art_blocks)
    total_items = sum(int(artifact_counts.get(name) or 0) for name, _labels in art_blocks)
    sub_lines = []
    for item in result.get("subtags") or []:
        labels = report_tags.get(item) or [item.replace(" -- ", " — ")]
        mapped = item.replace(" -- ", " — ")
        parts = []
        for label in labels:
            parts.append(f"{label}{_count_suffix(_lookup_count(tag_counts, label))}")
        if parts:
            sub_lines.append(f"{mapped}  ←  " + "; ".join(parts))
        else:
            sub_lines.append(mapped)
    unmatched_a = list(result.get("unmatched_artifacts") or [])
    unmatched_t = list(result.get("unmatched_tags") or [])
    found = len(result.get("raw_artifacts") or [])
    report_name = os.path.basename(result.get("report_path") or "AXIOM report")
    version = result.get("axiom_version") or ""
    kind = (result.get("source_kind") or "html").upper()

    win = tk.Toplevel(parent)
    win.title("AXIOM Report")
    apply_theme(win)
    win.transient(parent)
    win.minsize(520, 420)

    frame = ttk.Frame(win, padding=16)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(2, weight=3)
    frame.rowconfigure(4, weight=1)

    header = f"AXIOM {kind} report parsed"
    if version:
        header += f"  ·  v{version}"
    ttk.Label(frame, text=header, style="CategoryHeader.TLabel").grid(
        row=0, column=0, sticky="w"
    )
    ttk.Label(
        frame,
        text=(
            f"{report_name}   ·   {len(matched)} of {found} artifact types matched"
            + (f"   ·   {total_items} items" if total_items else "")
        ),
        style="Hint.TLabel",
    ).grid(row=1, column=0, sticky="w", pady=(2, 10))

    art_frame = ttk.LabelFrame(
        frame,
        text=(
            f"Matched artifacts ({len(matched)})  ·  imported tags ({imported_count})"
            + (f"  ·  {total_items} items" if total_items else "")
        ),
        padding=8,
    )
    art_frame.grid(row=2, column=0, sticky="nsew")
    art_frame.columnconfigure(0, weight=1)
    art_frame.rowconfigure(0, weight=1)
    art_text = tk.Text(
        art_frame,
        wrap="word",
        height=14,
        bg=COLORS["entry_bg"],
        fg=COLORS["text"],
        insertbackground=COLORS["text"],
        selectbackground=COLORS["accent_dark"],
        selectforeground=COLORS["button_fg"],
        highlightthickness=1,
        highlightbackground=COLORS["border"],
        relief="flat",
        font=("Segoe UI", 10),
        padx=8,
        pady=6,
        cursor="arrow",
    )
    art_scroll = ttk.Scrollbar(art_frame, orient="vertical", command=art_text.yview)
    art_text.grid(row=0, column=0, sticky="nsew")
    art_scroll.grid(row=0, column=1, sticky="ns")
    art_text.configure(yscrollcommand=art_scroll.set)
    art_text.tag_configure("artifact", font=("Segoe UI", 10, "bold"), spacing1=6, spacing3=2)
    art_text.tag_configure("tag", lmargin1=22, lmargin2=22, foreground=COLORS.get("hint", COLORS["text"]))
    art_text.tag_configure("empty", lmargin1=22, lmargin2=22, foreground=COLORS.get("hint", COLORS["text"]))
    if art_blocks:
        for name, labels in art_blocks:
            item_count = _lookup_count(artifact_counts, name)
            art_text.insert("end", f"{name}{_count_suffix(item_count)}\n", "artifact")
            if labels:
                for label in labels:
                    tag_count = _lookup_count(tag_counts, label)
                    art_text.insert("end", f"Tag: {label}{_count_suffix(tag_count)}\n", "tag")
            else:
                art_text.insert("end", "(no imported tags)\n", "empty")
    else:
        art_text.insert("end", "No catalog matches.")
    art_text.configure(state="disabled")

    tag_frame = ttk.LabelFrame(frame, text=f"Media tags ({len(sub_lines)})", padding=8)
    tag_frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))
    tag_frame.columnconfigure(0, weight=1)
    if sub_lines:
        for item in sub_lines:
            ttk.Label(tag_frame, text=f"•  {item}").pack(anchor="w", pady=1)
    else:
        ttk.Label(
            tag_frame,
            text="None mapped. Child Pornography, Child Erotica, and Age Difficult can still be checked by hand.",
            style="Hint.TLabel",
            wraplength=460,
        ).pack(anchor="w")

    extra_items = [f"Unmatched artifact: {name}" for name in unmatched_a]
    extra_items.extend(f"Custom tag: {name}" for name in unmatched_t)
    if extra_items:
        extra_frame = ttk.LabelFrame(frame, text="Needs review", padding=8)
        extra_frame.grid(row=4, column=0, sticky="nsew", pady=(10, 0))
        extra_frame.columnconfigure(0, weight=1)
        extra_frame.rowconfigure(0, weight=1)
        extra_list = tk.Listbox(
            extra_frame,
            height=min(6, max(3, len(extra_items))),
            activestyle="none",
            exportselection=False,
            bg=COLORS["entry_bg"],
            fg=COLORS["text"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            relief="flat",
            font=("Segoe UI", 10),
        )
        extra_scroll = ttk.Scrollbar(extra_frame, orient="vertical", command=extra_list.yview)
        extra_list.grid(row=0, column=0, sticky="nsew")
        extra_scroll.grid(row=0, column=1, sticky="ns")
        extra_list.configure(yscrollcommand=extra_scroll.set)
        _fill_listbox(extra_list, extra_items, "")
    else:
        frame.rowconfigure(4, weight=0)

    ttk.Label(
        frame,
        text="These items are pre-checked. Open Select Artifacts to confirm, add, or remove any of them.",
        style="Hint.TLabel",
        wraplength=480,
    ).grid(row=5, column=0, sticky="w", pady=(12, 8))

    ttk.Button(frame, text="OK", command=win.destroy, width=12).grid(row=6, column=0, sticky="e")

    win.update_idletasks()
    width = max(560, win.winfo_reqwidth())
    height = max(480, min(640, win.winfo_reqheight() + 20))
    try:
        px = parent.winfo_rootx() + max(0, (parent.winfo_width() - width) // 2)
        py = parent.winfo_rooty() + max(20, (parent.winfo_height() - height) // 4)
    except Exception:
        px, py = 80, 80
    win.geometry(f"{width}x{height}+{px}+{py}")
    win.grab_set()
    win.focus_set()
    parent.wait_window(win)


def infer_device_type(parent, device_class):
    hinted = getattr(parent, "_artifact_device_type", None)
    if hinted:
        return hinted
    if (device_class or "").lower() in {"computer", "pc", "computer/storage", "storage"}:
        return "Computer"
    if (device_class or "").lower() in {"warrant", "sw", "search warrant", "warrant return", "cloud"}:
        return "Cloud"
    extraction = getattr(parent, "extraction_file", "") or ""
    if extraction and hasattr(parent, "parse_extraction_file") and hasattr(parent, "determine_device_type"):
        try:
            data = parent.parse_extraction_file()
            return parent.determine_device_type(data)
        except Exception:
            return None
    return None


def load_axiom_report_path(parent, path, device_class="mobile"):
    from tkinter import messagebox
    if not path:
        return False
    target = os.path.normpath(path.strip().strip("{}"))
    if target.lower().endswith(".zip"):
        messagebox.showwarning(
            "AXIOM Report",
            "Point at the extracted export folder, Report.html, or Export.xml.\n"
            "Zipped exports need to be unpacked first.",
        )
        return False
    try:
        result = parse_axiom_report(
            target,
            device_class=device_class,
            device_type=infer_device_type(parent, device_class),
        )
    except Exception as exc:
        messagebox.showerror("AXIOM Report", str(exc))
        return False

    apply_parse_to_form(parent, result)
    display_name = os.path.basename(result["report_path"])
    label = getattr(parent, "axiom_report_drop_label", None)
    if label is not None:
        count = len(parent.selected_artifacts)
        label.configure(text=f"Selected: {display_name}  ({count} artifacts pre-checked)")
    show_parse_dialog(parent, result)
    return True


def clear_axiom_report(parent):
    parent.axiom_report_path = ""
    parent.axiom_report_parse = None
    parent.axiom_report_tags = {}
    parent.axiom_tag_counts = {}
    parent.axiom_version = ""
    label = getattr(parent, "axiom_report_drop_label", None)
    if label is not None:
        label.configure(text=DROP_HINT)


def attach_axiom_report_picker(parent, software_content, device_class="mobile"):
    """Add the optional Report.html drop zone under Select Artifacts."""
    from tkinter import ttk
    from tkinterdnd2 import DND_FILES
    from report_common import ask_open_file

    parent.axiom_report_path = getattr(parent, "axiom_report_path", "") or ""
    parent.axiom_report_parse = getattr(parent, "axiom_report_parse", None)
    parent.axiom_report_tags = getattr(parent, "axiom_report_tags", {}) or {}
    parent.axiom_tag_counts = getattr(parent, "axiom_tag_counts", {}) or {}
    parent.axiom_version = getattr(parent, "axiom_version", "") or ""
    parent.axiom_report_device_class = device_class

    label = ttk.Label(
        software_content,
        text=DROP_HINT,
        relief="solid",
        borderwidth=2,
        padding=8,
        wraplength=520,
        justify="left",
    )
    label.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(4, 2))
    label.grid_remove()
    parent.axiom_report_drop_label = label

    def browse(_event=None):
        path = ask_open_file(
            filetypes=[
                ("AXIOM report", "*.html *.htm *.xml"),
                ("HTML files", "*.html *.htm"),
                ("XML files", "*.xml"),
                ("All files", "*.*"),
            ],
            folder_kind="extraction",
            title="Select AXIOM Report.html or Export.xml",
        )
        if path:
            load_axiom_report_path(parent, path, device_class=device_class)

    def drop(event):
        try:
            dropped = list(parent.tk.splitlist(event.data))
        except Exception:
            dropped = [event.data]
        if dropped:
            load_axiom_report_path(parent, dropped[0], device_class=device_class)

    label.bind("<Button-1>", browse)
    try:
        label.drop_target_register(DND_FILES)
        label.dnd_bind("<<Drop>>", drop)
    except Exception:
        pass
    parent._browse_axiom_report = browse
    parent._drop_axiom_report = drop
    return label


def show_axiom_report_picker(parent):
    label = getattr(parent, "axiom_report_drop_label", None)
    if label is not None:
        label.grid()


def hide_axiom_report_picker(parent):
    label = getattr(parent, "axiom_report_drop_label", None)
    if label is not None:
        label.grid_remove()
    clear_axiom_report(parent)
