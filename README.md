# MFDR — Mega File Detector Renamer

Batch-renaming GUI for scanned documents, PDFs, and images. Extracts
titles and dates from PDF text layers, DOCX content, and OCR (tesseract)
to generate meaningful filenames, with a live preview so every rename
is human-approved before it happens.

Built to replace the Writer/Adobe/Excel/Gimp shuffle for organizing a
40,000+ file document archive.

## Features

- **Preview with zoom** — page-1 render for PDF, DOCX, JPG, PNG, BMP,
  TIFF (fit, 25%-400%), rendered in background threads
- **Resizable file list** — draggable divider for long filenames
- **Smart suggestions** — metadata extraction via `pdftotext`,
  `pdfinfo`, python-docx, and tesseract OCR for scans
- **Multi-select batch rename** — Ctrl/Shift-click a page sequence,
  approve once, files become `BaseName - 01.ext`, `- 02.ext`, ...
- **Duplicate detection** — content-hash (MD5) scan flags identical
  files with a `[DUP]` marker so redundant copies die in one click
- **Contextual Delete** — Delete key removes text while editing,
  deletes files (with confirmation) when the list has focus
- **Dry-run by default** — flip off when you trust a folder
- **Undo logs** — every session writes JSON; revert with
  `python3 mf_renamer.py --undo rename_log_<timestamp>.json`
- **Folder history** — last 20 folders remembered, one-click clear

## Requirements

    sudo apt install poppler-utils tesseract-ocr libreoffice
    pip3 install customtkinter Pillow

## Usage

    python3 mf_renamer_gui.py

Load a folder, review the preview and suggested name, edit if needed,
**Enter** to approve, **Tab** to skip, **Delete** to remove a file.

## Keyboard map

| Key | Action |
|-----|--------|
| Enter | Approve rename |
| Tab | Skip to next file |
| Delete | Delete file (or edit text, if focus is in a field) |
| +/-/Fit | Preview zoom |

## Project layout

    mf_renamer.py       core logic: metadata extraction, OCR, rename engine
    mf_renamer_gui.py   customtkinter/tkinter GUI

## License

MIT
