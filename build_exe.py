"""Build DFR Writer.exe with the Omega icon.

From this folder, in Command Prompt:

    build_exe
    python build_exe.py
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SPEC = HERE / "DFR_Writer.spec"
ICON = HERE / "assets" / "DFR_Writer.ico"
DIST = HERE / "dist"
EXE_NAME = "DFR Writer.exe"

# Same names PyInstaller rejects in PyInstaller.compat.check_requirements()
OBSOLETE_BACKPORTS = ("enum34", "typing", "pathlib")


def _distribution(name):
    try:
        from importlib.metadata import distribution
        return distribution(name)
    except Exception:
        return None


def _site_package_dirs():
    import site

    dirs = []
    try:
        dirs.extend(site.getsitepackages() or [])
    except Exception:
        pass
    try:
        user = site.getusersitepackages()
        if user:
            dirs.append(user)
    except Exception:
        pass
    for p in sys.path:
        if p and "site-packages" in p.replace("\\", "/").lower():
            dirs.append(p)
    out = []
    seen = set()
    for d in dirs:
        try:
            key = os.path.normcase(os.path.abspath(d))
        except Exception:
            continue
        if key not in seen and os.path.isdir(d):
            seen.add(key)
            out.append(Path(d))
    return out


def _purge_backport_files(name):
    """Delete leftover site-packages files if pip uninstall left metadata behind."""
    removed = False
    for site_dir in _site_package_dirs():
        candidates = [
            site_dir / name,
            site_dir / f"{name}.py",
            site_dir / f"{name}.pyc",
        ]
        for child in site_dir.glob(f"{name}-*"):
            if child.name.endswith(".dist-info") or child.name.endswith(".egg-info"):
                candidates.append(child)
        for path in candidates:
            if not path.exists():
                continue
            try:
                if path.is_dir():
                    shutil.rmtree(path)
                else:
                    path.unlink()
                print(f"    deleted {path}")
                removed = True
            except OSError as exc:
                print(f"    could not delete {path}: {exc}")
    return removed


def remove_obsolete_backports():
    """Uninstall old stdlib backports that make current PyInstaller refuse to start."""
    found = []
    for name in OBSOLETE_BACKPORTS:
        dist = _distribution(name)
        if dist is None:
            continue
        loc = ""
        try:
            loc = str(dist.locate_file(""))
        except Exception:
            loc = ""
        found.append((name, loc))

    if not found:
        return 0

    for name, loc in found:
        print(f"Removing obsolete '{name}' backport (incompatible with PyInstaller)...")
        if loc:
            print(f"    {loc}")
        code = subprocess.call(
            [sys.executable, "-m", "pip", "uninstall", "-y", name]
        )
        if code != 0 or _distribution(name) is not None:
            print(f"pip uninstall did not fully remove '{name}'. Cleaning leftover files...")
            _purge_backport_files(name)

        if _distribution(name) is not None:
            print()
            print(f"Still detected '{name}'. Close other Python windows and run:")
            print(f'    "{sys.executable}" -m pip uninstall -y {name}')
            return 1
    return 0


def ensure_pyinstaller():
    try:
        import PyInstaller  # noqa: F401
        return
    except Exception:
        pass
    print("Installing PyInstaller...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])


def copy_sidecar(src, dest_dir):
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / src.name
    if src.is_dir():
        if dest.exists():
            shutil.rmtree(dest)
        shutil.copytree(src, dest)
    elif src.exists():
        shutil.copy2(src, dest)


def main():
    os.chdir(HERE)
    if not SPEC.exists():
        print(f"Missing spec file: {SPEC}")
        return 1
    if not ICON.exists():
        print(f"Missing icon: {ICON}")
        return 1

    if remove_obsolete_backports() != 0:
        print("Build failed.")
        return 1

    ensure_pyinstaller()
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", str(SPEC)]
    print("Running:", " ".join(cmd))
    result = subprocess.call(cmd)
    if result != 0:
        print("Build failed.")
        return result

    exe = DIST / EXE_NAME
    if not exe.exists():
        matches = list(DIST.glob("*.exe"))
        if matches:
            exe = matches[0]
        else:
            print("Build finished but no .exe was found in dist\\")
            return 1

    # Official templates stay inside the exe. The first run asks where to
    # create the user "DFR Templates" folder and copies them there.
    copy_sidecar(HERE / "assets", DIST)

    print()
    print(f"Built: {exe}")
    print("Explorer icon is DFR_Writer.ico. The name under the icon is the exe filename.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
