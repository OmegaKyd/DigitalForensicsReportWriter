# PyInstaller spec for Digital Forensics Report Writer
# Builds "DFR Writer.exe" with the Omega icon (Explorer file icon).
# Run from this folder:  pyinstaller --noconfirm --clean DFR_Writer.spec

import os
from PyInstaller.utils.hooks import collect_all, collect_data_files

block_cipher = None
here = os.path.abspath(".")

datas = [
    (os.path.join(here, "assets"), "assets"),
]
settings_json = os.path.join(here, "settings.json")
if os.path.isfile(settings_json):
    datas.append((settings_json, "."))
templates = os.path.join(here, "Templates")
if os.path.isdir(templates):
    datas.append((templates, "Templates"))

hiddenimports = ["tkinterdnd2"]
binaries = []
try:
    dnd_datas, dnd_binaries, dnd_hidden = collect_all("tkinterdnd2")
    datas += dnd_datas
    binaries += dnd_binaries
    hiddenimports += dnd_hidden
except Exception:
    pass
try:
    datas += collect_data_files("tkinterdnd2")
except Exception:
    pass

a = Analysis(
    [os.path.join(here, "start_screen.py")],
    pathex=[here],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pytest", "unittest"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="DFR Writer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    icon=os.path.join(here, "assets", "DFR_Writer.ico"),
)
