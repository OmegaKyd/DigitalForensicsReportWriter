"""Write uncaught exceptions to a crash log the examiner can send."""
import os
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path


_HOOKED = False
_LAST_NOTICE = 0.0
_NOTICE_GAP = 4.0
LOG_NAME = "crash.log"
MAX_LOG_BYTES = 512 * 1024


def _app_meta():
    try:
        from ui_theme import APP_NAME, APP_VERSION
        return APP_NAME, APP_VERSION
    except Exception:
        return "Digital Forensics Report Writer", "unknown"


def logs_dir():
    """Same user-data folder as settings when frozen; script folder when running source."""
    app_name, _version = _app_meta()
    if getattr(sys, "frozen", False):
        if sys.platform == "win32":
            base = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        else:
            base = os.environ.get("XDG_CONFIG_HOME") or str(Path.home() / ".config")
        path = Path(base) / app_name / "logs"
    else:
        path = Path(__file__).resolve().parent / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def crash_log_path():
    return logs_dir() / LOG_NAME


def _rotate_if_needed(path):
    try:
        if path.is_file() and path.stat().st_size >= MAX_LOG_BYTES:
            bak = path.with_name("crash.old.log")
            if bak.exists():
                bak.unlink()
            path.replace(bak)
    except Exception:
        pass


def write_crash_report(exc_type, exc_value, exc_tb):
    """Append one crash block. Returns the log path, or None if writing failed."""
    if exc_type is None or issubclass(exc_type, KeyboardInterrupt):
        return None
    app_name, version = _app_meta()
    path = crash_log_path()
    _rotate_if_needed(path)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    kind = getattr(exc_type, "__name__", str(exc_type))
    lines = [
        "",
        "=" * 72,
        f"Time: {now}",
        f"App: {app_name} {version}",
        f"Frozen: {bool(getattr(sys, 'frozen', False))}",
        f"Python: {sys.version.replace(chr(10), ' ')}",
        f"Platform: {sys.platform}",
        f"Executable: {sys.executable}",
        f"Exception: {kind}: {exc_value}",
        "-" * 72,
        "".join(traceback.format_exception(exc_type, exc_value, exc_tb)).rstrip(),
        "=" * 72,
        "",
    ]
    try:
        with open(path, "a", encoding="utf-8") as handle:
            handle.write("\n".join(lines))
        return path
    except Exception:
        return None


def _show_notice(path, exc_value):
    global _LAST_NOTICE
    now = time.monotonic()
    if now - _LAST_NOTICE < _NOTICE_GAP:
        return
    _LAST_NOTICE = now
    text = (
        "The program hit an unexpected error.\n\n"
        f"{exc_value}\n\n"
        "A crash log was saved so it can be sent for troubleshooting:\n"
        f"{path}"
    )
    try:
        from tkinter import messagebox
        messagebox.showerror("Unexpected Error", text)
    except Exception:
        try:
            print(text, file=sys.stderr)
        except Exception:
            pass


def handle_exception(exc_type, exc_value, exc_tb):
    path = write_crash_report(exc_type, exc_value, exc_tb)
    if path is not None:
        _show_notice(path, exc_value)


def _tk_callback_exception(exc, val, tb):
    handle_exception(exc, val, tb)


def install_crash_logging(root=None):
    """Hook Python and Tk so uncaught errors go to crash.log."""
    global _HOOKED
    if not _HOOKED:
        sys.excepthook = handle_exception
        try:
            import threading
            threading.excepthook = lambda args: handle_exception(
                args.exc_type, args.exc_value, args.exc_traceback
            )
        except Exception:
            pass
        _HOOKED = True
    if root is not None:
        try:
            root.report_callback_exception = _tk_callback_exception
        except Exception:
            pass
    return crash_log_path()


def open_crash_logs(parent=None):
    """Open the logs folder, or the log file if the folder cannot be opened."""
    folder = logs_dir()
    path = crash_log_path()
    target = folder if folder.is_dir() else path
    try:
        if sys.platform == "win32":
            os.startfile(str(target))
        elif sys.platform == "darwin":
            os.spawnlp(os.P_NOWAIT, "open", "open", str(target))
        else:
            os.spawnlp(os.P_NOWAIT, "xdg-open", "xdg-open", str(target))
        return path
    except Exception as exc:
        from tkinter import messagebox
        messagebox.showinfo(
            "Crash Logs",
            f"Crash logs are saved at:\n{path}\n\n{exc}",
            parent=parent,
        )
        return path
