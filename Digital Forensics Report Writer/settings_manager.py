import json
import os
import sys
from pathlib import Path

from ui_theme import APP_NAME, LEGACY_APP_NAME, resource_dir


class SettingsManager:
    def __init__(self):
        # Factory copy baked into the exe (read-only when frozen).
        self.bundled_file = resource_dir() / "settings.json"
        # User changes: AppData when running the exe, settings.json beside
        # the scripts when running from source. Never write next to the .exe.
        if getattr(sys, "frozen", False):
            self.settings_file = self._appdata_dir() / "settings.json"
            self.legacy_file = Path(sys.executable).resolve().parent / "settings.json"
            self.legacy_appdata_file = self._appdata_dir(LEGACY_APP_NAME, create=False) / "settings.json"
        else:
            self.settings_file = Path(__file__).resolve().parent / "settings.json"
            self.legacy_file = None
            self.legacy_appdata_file = None

        self.default_settings = {
            "examiner_agency_type": "",
            "examiner_agency_custom": "",
            "examiner_title": "",
            "examiner_name": "",
            "dfr_number_prefix": "DFR2026-",
            "warrant_report_prefix": "WDR2026-",
            "default_service_provider": "",
            "last_template_dir": "",
            "last_extraction_dir": "",
            "last_export_dir": "",
            "dfr_templates_dir": "",
            "remembered_request_titles": [],
            "remembered_request_agencies": [],
            "version": "1.0.1",
        }

    def _appdata_dir(self, app_name=None, create=True):
        if sys.platform == "win32":
            base = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        else:
            base = os.environ.get("XDG_CONFIG_HOME") or str(Path.home() / ".config")
        path = Path(base) / (app_name or APP_NAME)
        if create:
            path.mkdir(parents=True, exist_ok=True)
        return path

    def _read_json(self, path):
        try:
            if path and Path(path).is_file():
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception as e:
            print(f"Error reading settings {path}: {e}")
        return {}

    def load_settings(self):
        settings = self.default_settings.copy()
        # Baked-in copy from the exe (or the project file when running source).
        if self.bundled_file != self.settings_file:
            settings.update(self._read_json(self.bundled_file))
        # One-time migrate of an old settings.json that sat next to the exe.
        if self.legacy_file and self.legacy_file != self.settings_file:
            settings.update(self._read_json(self.legacy_file))
        # v1.0.0 and earlier stored AppData under the previous product name.
        if (
            self.legacy_appdata_file
            and self.legacy_appdata_file != self.settings_file
            and not self.settings_file.is_file()
        ):
            settings.update(self._read_json(self.legacy_appdata_file))
        settings.update(self._read_json(self.settings_file))
        return settings

    def save_settings(self, settings):
        try:
            merged = self.load_settings()
            merged.update(settings)
            self.settings_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(merged, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error saving settings: {e}")
            return False

    def get_settings_file_path(self):
        return str(self.settings_file)

# Ω Digital Forensics Report Writer Ω (ver. 1.0.1) © 2026 #
