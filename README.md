# Digital Forensics Report Writer

**Version 1.1.0** (changes since v1.0.6-beta.1; cumulative since v1.0.5)

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

Official base templates ship in the project `Templates/` folder as `DFR Mobile.docx`, `DFR Computer.docx`, `DFR Storage.docx`, and `DFR SW Return.docx`. They are generic, not agency-specific. On first run they are copied into a writable **DFR Templates** folder. Customize copies there rather than the files in `Templates/`.

## Features

- Optional Notes tab; notes print at `PY_NOTES`
- Mobile Device Prep tab Android / Apple iOS prep checklists check tagged boxes on the official base template (`PY_ANDROID_*`, `PY_IOS_*`)
- **No Extraction Completed** next to No Evidence Found; skips required fields except the template and writes only completed values
- File → Save Progress / Open Unfinished Reports to pause and resume a form; File → Manage Unfinished Reports… deletes drafts without generating; generating a report also deletes that draft
- Tools → Report Number Prefix… sets the default value filled into the Report Number field
- Shared dark forensic theme and Omega header across the start screen and report windows
- Examiner name, title, and agency remembered between sessions
- Shared, editable officer-title list for requesting officer, examiner, and transfer officer (**Tools → Edit Officer Titles…**)
- Requesting Officer is a remembered combobox with the same typeahead as Agency and Title (**Tools → Edit Requesting Officers…**)
- Warrant Data Returns Service Provider field remembers previous names and typeahead-completes them (**Tools → Edit Service Providers…**)
- Editable canned narrative for every report type (**Tools → Edit Paragraphs…**); factory wording stays in the program, user overrides are stored separately
- Calendar date pickers on report date fields
- Case Agent can fill requesting-officer fields from the examiner so `PY_REQOFF` / `PY_REQAGENCY` still flow through the existing placeholder path
- Drag-and-drop and file pickers for extraction logs and PDFs
- Cellebrite Summary Report / Quick View parsing (`cellebrite_pdf.py`)
- GrayKey Progress Report software version fills `PY_GKVER` on DFR Mobile
- Computer modules also accept Cellebrite Digital Collector acquisition logs (`PY_DCVER`)
- First-listed identifier rule for IMEI, ICCID, and carrier lists
- GUI values win over parsed values when both exist
- Placeholder preview before a report is written
- Suggested output names from DFR number, report type, owner, and model
- Tools menu to add, rename, remove, or relocate templates
- **Tools → Manage Templates → Restore Official Templates** copies missing packaged `.docx` files back; deleted working copies stay deleted until you restore them
- Magnet Axiom artifact picker on Full Mobile Exam, Full Computer Exam, and Warrant Data Returns (`magnet_artifacts.py` / `magnet_artifacts.json`)
- Optional AXIOM HTML or XML tagged-export drop when Axiom is checked (`axiom_html_report.py`): pre-checks catalog artifacts and Pictures/Videos CP / Child Erotica / Age Difficult boxes from examiner tags
- Confirmation list shows unique item counts per artifact type and per imported tag; dual-tagged items count once toward the artifact and once on each tag
- Generated artifact headings are bold and list-indented; each examiner tag gets `Tag:`, `This tag contains N artifacts.`, and `(DEVICE ARTIFACTS)`
- `PY_EXAMINEVER` is filled with the Magnet AXIOM version from the loaded HTML/XML export (`PY_EXAMINE` still works on older templates)
- Warrant Data Returns uses the same tabbed form layout as the other report windows
- Artifact list is limited by device class (mobile vs computer/storage vs warrant/cloud); Cloud and Refined Results appear on mobile and computer classes
- Help → About includes a clickable GitHub link
- Help menu listing every supported `PY_` token
- Uncaught errors write a crash log; **Help → Open Crash Logs…** opens the folder so the file can be sent for troubleshooting

## What's new in v1.1.0

Compared with v1.0.6-beta.1:

- Requesting Officer is a remembered typeahead field (**Tools → Edit Requesting Officers…**)
- Form tabs span the full window; mobile Device Prep sits between Device Info and Output
- Crash log with **Help → Open Crash Logs…**
- **No Extraction Completed** generates a report from whatever is already filled
- GrayKey Progress Report OS version fills `PY_GKVER`
- Storage templates no longer warn about computer-only tags
- Computer Portable Device Info completion and `PY_ACQUIRE` fill

Tagged templates, `PY_NOTES`, authority/processing checkboxes, FTK 8.3 parsing, and `PY_EXAMINEVER` from v1.0.6-beta.1 remain.

Full detail is in the project-root `RELEASE_NOTES.md`.

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
- **Tools → Edit Requesting Officers…** edits the remembered officer-name list. Typing a name that is not listed and then generating a report adds it. The list stays alphabetical.
- **Tools → Edit Service Providers…** edits the Warrant Data Returns provider list. New names are added when a report is generated. The list stays alphabetical.

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
├── crash_log.py               # uncaught-exception log
├── app_menu.py                # File / Tools / Help
├── drafts_manager.py          # notes, checklists, unfinished reports
├── ui_theme.py                # theme, APP_NAME, APP_VERSION, header bar
├── settings_manager.py        # load / save settings
├── report_common.py           # placeholders, templates, titles, preview, naming
├── paragraphs_defaults.py     # factory narrative
├── paragraphs_manager.py      # load / save / fill narrative
├── paragraphs_editor.py       # Tools → Edit Paragraphs
├── titles_editor.py           # Tools → Edit Officer Titles
├── agencies_editor.py         # Tools → Edit Agencies
├── officers_editor.py         # Tools → Edit Requesting Officers
├── providers_editor.py        # Tools → Edit Service Providers
├── magnet_artifacts.py        # Magnet Axiom artifact picker
├── magnet_artifacts.json      # bundled artifact catalog
├── axiom_html_report.py       # AXIOM HTML/XML tagged-export parser
├── cellebrite_pdf.py          # Cellebrite PDF import
├── mobile_*.py / pc_*.py / sw_data_review.py
├── Templates/                 # official DFR Word templates (do not edit here)
├── assets/                    # Omega.png, icons
├── build_exe.py / DFR_Writer.spec
├── LICENSE
└── README.md
```
Release notes live in the project root `RELEASE_NOTES.md`.

## License

MIT. See [LICENSE](LICENSE).
