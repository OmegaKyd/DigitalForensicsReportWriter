# Digital Forensics Report Writer

**Version 1.0.1**

Desktop application for writing digital forensic reports from mobile extractions, computer acquisitions, and search-warrant data returns. The examiner fills case fields (or imports them from vendor reports), chooses a Word template, previews the `PY_` placeholders, and writes a completed `.docx` report.

Formerly *Digital Evidence Report Writer*. Windows is the supported platform.

GitHub: [https://github.com/omegakyd](https://github.com/omegakyd)

## Report types

| Start-screen button | Module | Typical sources | Official template |
| --- | --- | --- | --- |
| Mobile — Portable Case | `mobile_portable_case.py` | Cellebrite UFD / Summary / Quick View, GrayKey PDF | `DFR Storage.docx` |
| Mobile — Full Exam | `mobile_full_exam.py` | Cellebrite UFD / Summary / Quick View, GrayKey PDF | `DFR Mobile.docx` |
| PC — Portable Case | `pc_portable_case.py` | TX1, FTK Imager, X-Ways, or Cellebrite Digital Collector log | `DFR Storage.docx` |
| PC — Full Exam | `pc_full_exam.py` | TX1, FTK Imager, X-Ways, or Cellebrite Digital Collector log | `DFR Computer.docx` |
| Warrant Data Returns | `sw_data_review.py` | Warrant, subpoena, or service-provider return | `DFR SW Return.docx` |

Official templates ship in the project `Templates/` folder and are copied into a writable **DFR Templates** folder on first run. Do not edit the files in `Templates/` if you are working from this project tree; customize copies in **DFR Templates** instead.

## Features

- Shared dark forensic theme and Omega header across the start screen and report windows
- Examiner name, title, and agency remembered between sessions
- Shared, editable officer-title list for requesting officer, examiner, and transfer officer (**Tools → Edit Officer Titles…**)
- Editable canned narrative for every report type (**Tools → Edit Paragraphs…**); factory wording stays in the program, user overrides are stored separately
- Calendar date pickers on report date fields
- Case Agent can fill requesting-officer fields from the examiner so `PY_REQOFF` / `PY_REQAGENCY` still flow through the existing placeholder path
- Drag-and-drop and file pickers for extraction logs and PDFs
- Cellebrite Summary Report / Quick View parsing (`cellebrite_pdf.py`)
- Computer modules also accept Cellebrite Digital Collector acquisition logs (`PY_DCVER`)
- First-listed identifier rule for IMEI, ICCID, and carrier lists
- GUI values win over parsed values when both exist
- Placeholder preview before a report is written
- Suggested output names from DFR number, report type, owner, and model
- Tools menu to add, rename, remove, or relocate templates
- Help → About includes a clickable GitHub link
- Help menu listing every supported `PY_` token

## Requirements

Python 3 on Windows, plus the packages in `Requirements.txt`:

```
PyPDF2
lxml
python-docx
tkinterdnd2
Pillow
```

Install with:

```bat
py -3 -m pip install -r Requirements.txt
```

## Run from source

From the program folder:

```bat
Launch.bat
```

or:

```bat
py -3 start_screen.py
```

## Settings and user data

| How you run it | Where settings and paragraph overrides live |
| --- | --- |
| Source (`start_screen.py`) | `settings.json` and `paragraphs.json` next to the scripts |
| Frozen exe (`DFR Writer.exe`) | `%APPDATA%\Digital Forensics Report Writer\` |

Factory defaults are taken from the project `settings.json` and `paragraphs_defaults.py` and baked into the exe at build time. The exe does not write settings next to itself.

If you are upgrading from v1.0.0 (when the program was still named Digital Evidence Report Writer), the first run of the exe copies existing `settings.json` / `paragraphs.json` from `%APPDATA%\Digital Evidence Report Writer\` into the new folder. Your DFR Templates path is kept if it was already saved.

## Templates

1. First launch asks where to create **DFR Templates**.
2. Official `.docx` files are copied there if they are missing.
3. **Tools → Manage Templates…** adds, renames, or removes files in that folder.
4. **Tools → Change Template Folder Location…** moves the working folder later.

The project `Templates/` directory remains the official source copy.

## Narrative and titles

- **Tools → Edit Paragraphs…** edits the canned report wording. Tokens such as `{Request_Date}` are filled from the form. `PY_` tokens belong in Word templates, not in these paragraphs.
- **Tools → Edit Officer Titles…** edits the shared title list. Typing a title that is not listed and then saving a report adds it. Drag rows or use Move Up / Move Down to reorder. Revert restores the factory list.

## Build a Windows exe

From the same folder:

```bat
build_exe.bat
```

or:

```bat
py -3 build_exe.py
```

`build_exe.py` removes obsolete stdlib backports (`pathlib`, `enum34`, `typing`) that current PyInstaller will refuse, then builds with `DFR_Writer.spec`. Output is `dist\DFR Writer.exe` with the Omega icon from `assets\DFR_Writer.ico`.

## Project layout

```
Digital Forensics Report Writer/
├── start_screen.py            # launcher
├── app_menu.py                # File / Tools / Help
├── ui_theme.py                # theme, APP_NAME, APP_VERSION, header bar
├── settings_manager.py        # load / save settings
├── report_common.py           # placeholders, templates, titles, preview, naming
├── paragraphs_defaults.py     # factory narrative
├── paragraphs_manager.py      # load / save / fill narrative
├── paragraphs_editor.py       # Tools → Edit Paragraphs
├── titles_editor.py           # Tools → Edit Officer Titles
├── cellebrite_pdf.py          # Cellebrite PDF import
├── mobile_*.py / pc_*.py / sw_data_review.py
├── Templates/                 # official DFR Word templates (do not edit here)
├── assets/                    # Omega.png, icons
├── build_exe.py / DFR_Writer.spec
├── LICENSE
├── README.md
└── RELEASE_NOTES.md
```

## License

MIT. See [LICENSE](LICENSE).
