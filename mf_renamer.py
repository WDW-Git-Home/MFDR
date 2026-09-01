#!/usr/bin/env python3
"""
Mega File Detector Renamer (MFDR)
Suggests names for PDFs, images, and Office docs using embedded
metadata, text layers, and OCR. Renames only after human approval.

  python3 mf_renamer.py --dir <folder> --dry-run
  python3 mf_renamer.py --dir <folder> --recursive
  python3 mf_renamer.py --undo <rename_log_XXXX.json>
"""

import os
import re
import json
import sys
import argparse
import subprocess
import zipfile
import tempfile
from datetime import datetime

# ---------------------------------------------------------------- constants

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".webp"}
PROCESSABLE = {".pdf", ".docx"} | IMAGE_EXTS

NOISE = re.compile(
    r"^(\s*page\s+\d+.*|\d+\s*/\s*\d+|confidential|cc[-_ ]?\w+|"
    r"printed|created with|microsoft)", re.I)
URLISH = re.compile(r"(www\.|https?://|\.(com|net|org|edu|gov)\b)", re.I)
JUNK = re.compile(r"(?i)^(microsoft word|untitled|document ?\d*|scan ?\d*)")

MONTHS = {m: i + 1 for i, m in enumerate(
    ["january", "february", "march", "april", "may", "june", "july",
     "august", "september", "october", "november", "december"])}

# ------------------------------------------------------------------- utils

def run(cmd, timeout=60):
    """Run a subprocess, return stdout, or '' on any failure."""
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        return r.stdout or ""
    except Exception:
        return ""

def clean(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", str(name))
    return re.sub(r"\s+", " ", name).strip(" ._-")[:100] or "Untitled"

def is_title_candidate(s):
    s = s.strip()
    if not (5 <= len(s) <= 90):
        return False
    if NOISE.match(s) or JUNK.match(s):
        return False
    if URLISH.search(s):
        return False
    letters = sum(c.isalpha() for c in s)
    return letters >= 5 and letters / max(len(s), 1) > 0.5

def pick_title(lines):
    for s in lines:
        if is_title_candidate(s):
            return s.strip()
    return None

def find_date(text):
    """Extract a real date from document text. None if absent."""
    for line in text.splitlines():
        s = line.strip()
        if not (8 <= len(s) <= 90) or URLISH.search(s):
            continue
        m = re.search(r"\b(\d{4})-(\d{2})-(\d{2})\b", s)
        if m:
            return m.group(0)
        m = re.search(r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
                      r"[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})", s, re.I)
        if m:
            mon = MONTHS.get(m.group(1)[:3].lower())
            if mon:
                return f"{m.group(3)}-{mon:02d}-{int(m.group(2)):02d}"
        m = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", s)
        if m:
            return f"{m.group(3)}-{int(m.group(1)):02d}-{int(m.group(2)):02d}"
    return None

def mtime_date(path):
    return datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d")

def old_name_title(path):
    base = os.path.splitext(os.path.basename(path))[0]
    return re.sub(r"[_]+", " ", base).strip() or "Untitled"


def find_year(text):
    """Year-only fallback when no full date exists (brochures, catalogs)."""
    m = re.search(r"\b(19\d{2}|20\d{2})\b", text)
    return m.group(0) if m else None

def build_name(subject, date, pages, ext):
    """Subject - Date - Np.ext  (segments omitted when unknown)"""
    name = clean(subject)
    if date:
        name += f" - {date}"
    if pages and pages > 1:
        name += f" - {pages}p"
    return name + ext

# --------------------------------------------------------------------- PDF

def analyze_pdf(path):
    info = run(["pdfinfo", path], timeout=15)
    m = re.search(r"Pages:\s+(\d+)", info)
    pages = int(m.group(1)) if m else 0

    title, date, method = None, None, "fallback"

    # 1) Embedded metadata title
    tm = re.search(r"^Title:\s*(.+)$", info, re.MULTILINE)
    if tm and is_title_candidate(tm.group(1)):
        title, method = tm.group(1).strip(), "pdf-title"

    # 2) Embedded text layer, 3) OCR fallback
    text = ""
    if title is None:
        text = run(["pdftotext", "-l", "2", path, "-"], timeout=30).strip()
        if len(text) >= 40:
            method = "text"
        else:
            method = "OCR"
            with tempfile.TemporaryDirectory() as td:
                subprocess.run(
                    ["pdftoppm", "-f", "1", "-l", "1", "-r", "200", "-png",
                     path, os.path.join(td, "pg")],
                    capture_output=True, timeout=120)
                pngs = sorted(f for f in os.listdir(td) if f.startswith("pg"))
                if pngs:
                    text = run(["tesseract", os.path.join(td, pngs[0]), "-"],
                               timeout=90)
        title = pick_title(text.splitlines())

    date = find_date(text)

    if title is None:
        title = old_name_title(path)
    date = date or find_year(text)

    meta = {"method": method, "pages": pages, "date_found": date is not None}
    return build_name(title, date, pages, ".pdf"), meta

# ------------------------------------------------------------------ images

def analyze_image(path):
    ext = os.path.splitext(path)[1].lower()
    text = run(["tesseract", path, "-", "--psm", "3"], timeout=90)
    method = "OCR" if len(text.strip()) >= 20 else "no-text"
    title = pick_title(text.splitlines()) or old_name_title(path)
    date = find_date(text) or find_year(text)
    return build_name(title, date, None, ext), {"method": method}

# ------------------------------------------------------------- Office docs

def analyze_docx(path):
    try:
        with zipfile.ZipFile(path) as z:
            core = z.read("docProps/core.xml").decode("utf-8", "ignore")
    except Exception:
        return None
    title = None
    m = re.search(r"<dc:title>(.*?)</dc:title>", core, re.S)
    if m and is_title_candidate(m.group(1)):
        title = m.group(1).strip()
    m = re.search(r"<dcterms:created[^>]*>(\d{4}-\d{2}-\d{2})", core)
    date = m.group(1) if m else None
    if not title:
        title = old_name_title(path)
    return build_name(title, date, None, ".docx")

def analyze_any(path):
    """Single entry point for CLI and GUI."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return analyze_pdf(path)
    if ext in IMAGE_EXTS:
        return analyze_image(path)
    if ext == ".docx":
        name = analyze_docx(path)
        return (name, {"method": "docx"}) if name else (None, {"method": "error"})
    raise ValueError(f"Unsupported type: {ext}")

# ------------------------------------------------------------------- main

def collect_targets(folder, recursive):
    if recursive:
        found = []
        for root, dirs, names in os.walk(folder):
            dirs[:] = [d for d in dirs if not d.startswith("_")]
            for n in names:
                if os.path.splitext(n)[1].lower() in PROCESSABLE:
                    found.append(os.path.join(root, n))
        return sorted(found)
    return sorted(
        os.path.join(folder, n) for n in os.listdir(folder)
        if os.path.isfile(os.path.join(folder, n))
        and os.path.splitext(n)[1].lower() in PROCESSABLE)

def main():
    ap = argparse.ArgumentParser(description="Mega File Detector Renamer")
    ap.add_argument("--dir", required=True, help="Folder to process")
    ap.add_argument("--recursive", action="store_true", help="Include subfolders")
    ap.add_argument("--dry-run", action="store_true",
                    help="Preview suggestions, rename nothing")
    ap.add_argument("--undo", metavar="LOG", help="Revert a rename log")
    args = ap.parse_args()

    if args.undo:
        with open(args.undo) as f:
            entries = json.load(f)
        n = 0
        for e in entries:
            if os.path.exists(e["to"]) and not os.path.exists(e["from"]):
                os.rename(e["to"], e["from"])
                n += 1
        print(f"Restored {n} names.")
        sys.exit(0)

    targets = collect_targets(args.dir, args.recursive)
    if not targets:
        print(f"No processable files in {args.dir}")
        sys.exit(1)

    print(f"MFDR: {len(targets)} files to review\n")
    log = []
    renamed = 0

    for i, path in enumerate(targets, 1):
        print("=" * 72)
        print(f"[{i}/{len(targets)}] {os.path.basename(path)}")
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext == ".pdf":
                proposed, meta = analyze_pdf(path)
                print(f"  ({meta['method']}, {meta['pages']} pages)")
            elif ext in IMAGE_EXTS:
                proposed, meta = analyze_image(path)
                print(f"  ({meta['method']})")
            else:  # .docx
                proposed = analyze_docx(path)
                if proposed is None:
                    print("  !! unreadable metadata, skipped")
                    continue
                print("  (docx metadata)")
        except Exception as e:
            print(f"  !! extraction failed: {e}")
            continue

        print(f"  SUGGESTED: {os.path.basename(proposed)}")
        try:
            ans = input("  [Enter]=approve  e=edit  s=skip  q=quit : ").strip().lower()
        except (KeyboardInterrupt, EOFError):
            print("\nInterrupted — stopping.")
            break

        if ans == "q":
            break
        if ans == "s":
            continue
        if ans == "e":
            typed = input("  New name (no extension): ").strip()
            if not typed:
                print("  (skipped)")
                continue
            ext = os.path.splitext(path)[1]
            proposed = clean(typed) + ext

        dest = os.path.join(os.path.dirname(path), os.path.basename(proposed))
        base, ext = os.path.splitext(dest)
        n = 1
        while os.path.exists(dest):
            dest = f"{base} ({n}){ext}"
            n += 1

        if not args.dry_run:
            os.rename(path, dest)
        log.append({"from": path, "to": dest})
        print(f"  -> {os.path.basename(dest)}")

    print("=" * 72)
    if log and not args.dry_run:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        lp = os.path.join(args.dir, f"rename_log_{ts}.json")
        with open(lp, "w") as f:
            json.dump(log, f, indent=2)
        print(f"Renamed {len(log)} files. Log: {lp}")
        print(f"Undo: python3 {os.path.abspath(__file__)} --undo {lp}")
    elif log:
        print(f"Dry run complete: {len(log)} files would be renamed.")

if __name__ == "__main__":
    main()
