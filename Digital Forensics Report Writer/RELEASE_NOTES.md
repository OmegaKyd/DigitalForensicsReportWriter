# Release Notes — Digital Forensics Report Writer v1.0.1

**Release date:** 30 August 2026  
**Previous release:** Digital Evidence Report Writer v1.0.0

v1.0.1 renames the product and ships the narrative, title, date-picker, and Digital Collector work that landed after v1.0.0.

## Product name

The program is now **Digital Forensics Report Writer**.

Window titles, the header bar, Help → About, file footers, settings schema, README, and this license set all use the new name and **v1.0.1**.

The frozen exe still writes user files under `%APPDATA%\<product name>\`. Because the product name changed, v1.0.1 looks in the old folder first:

`%APPDATA%\Digital Evidence Report Writer\`

and copies `settings.json` / `paragraphs.json` into:

`%APPDATA%\Digital Forensics Report Writer\`

when the new files are not there yet. DFR Templates stays where you already pointed it.

Zipped historical copies that still say *Digital Evidence Report Writer* or `v1.0.0` / `v0.93` in the file name are storage-only backups and were not renamed.

Official Word templates in `Templates/` were not edited for this release.

## Highlights

- Product rename to Digital Forensics Report Writer, version badge **v1.0.1**
- **Tools → Edit Paragraphs…** — canned narrative is editable per report type; factory text stays in `paragraphs_defaults.py`
- **Tools → Edit Officer Titles…** — shared title list for requesting officer, examiner, and transfer officer
- Calendar date pickers on report date fields
- Cellebrite Digital Collector acquisition logs on PC Portable Case and PC Full Exam (`PY_DCVER`)
- Help → About still includes the clickable GitHub link [https://github.com/omegakyd](https://github.com/omegakyd)

## What changed since v1.0.0

### Narrative editor

- Factory paragraphs for Mobile Full, Mobile Portable, PC Full, PC Portable, and Warrant Data Returns
- User overrides stored in `paragraphs.json` (AppData when frozen)
- Field tokens such as `{Request_Date}` are filled from the form at generate time
- Revert one paragraph, one report type, or everything back to factory wording
- Open report windows refresh after a save

### Officer titles

- Shared list used by Requesting Officer Title, Examiner Title, and Transfer Officer Title
- Saving a report remembers a newly typed title
- Editor supports add, rename, remove, drag-reorder, move up/down, and revert to factory

### Computer acquisitions

- TX1, FTK Imager, X-Ways, **and** Cellebrite Digital Collector logs
- Digital Collector version maps to `PY_DCVER`

### Application

- Header badge and About box read `ui_theme.APP_VERSION` (`1.0.1`)
- Settings schema field `version` is `1.0.1`
- AppData folder follows the new product name, with a one-time migrate from the old name

## Requirements

See `Requirements.txt`: PyPDF2, lxml, python-docx, tkinterdnd2, Pillow. Python 3 on Windows.

## Known limits

- Designed and tested as a Windows desktop app (tkinter + optional frozen exe)
- Official Word templates in `Templates/` are source copies; customize files in DFR Templates instead
- PDF and log import depend on the layout of the vendor report. Unrecognized files can be classified manually when prompted
- Paragraph `{tokens}` are not the same as Word `PY_` placeholders. Use Help → Template Placeholders for template tokens, and Tools → Edit Paragraphs for narrative tokens

## Upgrade notes from v1.0.0

1. Replace the previous scripts or exe with the v1.0.1 tree / build.
2. First launch of the renamed exe migrates AppData from `Digital Evidence Report Writer` into `Digital Forensics Report Writer` if the new files do not exist yet.
3. Existing DFR Templates folder path is kept.
4. Keep using your working copies in DFR Templates. Do not edit the official files in `Templates/`.
5. If you customized narrative in an older copy that was not stored as `paragraphs.json`, re-enter those edits under Tools → Edit Paragraphs.
