#!/usr/bin/env python3
"""
MFDR GUI — Mega File Detector Renamer
Features: PDF/docx/images with metadata+OCR naming, page-1 preview with
zoom, resizable file list, multi-select batch rename, delete with
confirmation, duplicate detection (content hash), folder history,
dry-run default, undo logs.

  python3 mf_renamer_gui.py
"""

import os
import sys
import json
import hashlib
import threading
import subprocess
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

try:
    import customtkinter as ctk
    USE_CTK = True
    ctk.set_appearance_mode("dark")
except ImportError:
    USE_CTK = False

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mf_renamer as mfr

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "mfdr_settings.json")

def load_settings():
    try:
        with open(SETTINGS_FILE) as f:
            s = json.load(f)
    except Exception:
        s = {}
    s.setdefault("folders", [])
    return s

def save_settings(data):
    try:
        with open(SETTINGS_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass

def render_preview_png(path):
    """Render page 1 to a PNG path. PDF via pdftoppm, docx via
    LibreOffice, raster images via Pillow. None on failure."""
    ext = os.path.splitext(path)[1].lower()
    td = tempfile.mkdtemp(prefix="mfdr_")
    try:
        if ext == ".docx":
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf",
                 "--outdir", td, path],
                capture_output=True, timeout=90)
            pdfs = [f for f in os.listdir(td) if f.endswith(".pdf")]
            if not pdfs:
                return None
            path = os.path.join(td, pdfs[0])
            ext = ".pdf"
        if ext == ".pdf":
            subprocess.run(
                ["pdftoppm", "-f", "1", "-l", "1", "-r", "75", "-png",
                 path, os.path.join(td, "pg")],
                capture_output=True, timeout=60)
            pngs = sorted(f for f in os.listdir(td) if f.endswith(".png"))
            return os.path.join(td, pngs[0]) if pngs else None
        if ext in (".png", ".gif", ".ppm"):
            return path
        from PIL import Image
        out = os.path.join(td, "converted.png")
        with Image.open(path) as im:
            im.convert("RGB").save(out, "PNG")
        return out
    except Exception:
        import traceback
        traceback.print_exc()
        return None

class App:
    def __init__(self, root):
        self.root = root
        root.title("MFDR — Mega File Detector Renamer")
        root.geometry("1200x780")
        root.minsize(950, 620)

        self.targets = []
        self.current_path = None
        self.preview_img = None
        self.approved_log = []
        self.dup_paths = set()
        self.zoom_pct = 100
        self._preview_png = None
        self.settings = load_settings()

        self._build_top()
        self._build_middle()
        self._build_bottom()
        self._bind_keys()

    # ---- widget factories ----

    def _frame(self, parent, **kw):
        return ctk.CTkFrame(parent, **kw) if USE_CTK else tk.Frame(parent, **kw)

    def _button(self, parent, text, cmd, width=110):
        if USE_CTK:
            return ctk.CTkButton(parent, text=text, command=cmd, width=width)
        return tk.Button(parent, text=text, command=cmd, width=int(width / 8))

    def _label(self, parent, text):
        return ctk.CTkLabel(parent, text=text) if USE_CTK \
            else tk.Label(parent, text=text)

    def _entry(self, parent, var, width=400):
        if USE_CTK:
            return ctk.CTkEntry(parent, width=width, textvariable=var)
        return tk.Entry(parent, width=int(width / 7), textvariable=var)

    def _checkbox(self, parent, text, var):
        if USE_CTK:
            return ctk.CTkCheckBox(parent, text=text, variable=var)
        return tk.Checkbutton(parent, text=text, variable=var)

    # ------------------------------------------------------------ layout

    def _build_top(self):
        top = self._frame(self.root)
        top.pack(fill="x", padx=8, pady=(8, 4))

        self._label(top, "Folder:").pack(side="left", padx=(4, 2))
        self.var_dir = tk.StringVar(
            value=self.settings.get("last_folder",
                                    os.path.expanduser("~/Documents")))
        if USE_CTK:
            self.cmb_dir = ctk.CTkComboBox(
                top, variable=self.var_dir, width=420,
                values=self.settings.get("folders", []))
        else:
            self.cmb_dir = ttk.Combobox(
                top, textvariable=self.var_dir, width=65,
                values=self.settings.get("folders", []))
        self.cmb_dir.pack(side="left", padx=4)

        self._button(top, "Browse…", self.browse, 90).pack(side="left", padx=4)
        self._button(top, "Clear History", self.clear_history, 110).pack(
            side="left", padx=4)

        self.var_recursive = tk.BooleanVar(
            value=self.settings.get("recursive", False))
        self._checkbox(top, "Recursive", self.var_recursive).pack(
            side="left", padx=8)

        self.var_dry_run = tk.BooleanVar(
            value=self.settings.get("dry_run", True))
        self._checkbox(top, "Dry run", self.var_dry_run).pack(
            side="left", padx=8)

        self._button(top, "Load", self.load_folder, 90).pack(
            side="left", padx=4)
        self._button(top, "Save Log", self.write_log, 100).pack(
            side="left", padx=4)

    def _build_middle(self):
        paned = ttk.Panedwindow(self.root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=8, pady=4)

        left = tk.Frame(paned)
        paned.add(left, weight=1)

        self.lst = tk.Listbox(left, width=46, font=("Sans", 11),
                             selectmode=tk.EXTENDED, exportselection=False)
        self.lst.pack(side="left", fill="both", expand=True)
        self.lst.bind("<<ListboxSelect>>", self.on_select)

        sb = ttk.Scrollbar(left, command=self.lst.yview)
        sb.pack(side="right", fill="y")
        self.lst.config(yscrollcommand=sb.set)

        right = tk.Frame(paned)
        paned.add(right, weight=3)

        self.lbl_status = tk.Label(right, text="Load a folder to begin.",
                                   font=("Sans", 11), anchor="w")
        self.lbl_status.pack(fill="x", pady=(0, 4))

        zoombar = tk.Frame(right)
        zoombar.pack(fill="x", pady=(0, 2))
        tk.Button(zoombar, text="-", width=3,
                  command=self.zoom_out).pack(side="left", padx=2)
        tk.Button(zoombar, text="+", width=3,
                  command=self.zoom_in).pack(side="left", padx=2)
        tk.Button(zoombar, text="Fit", width=5,
                  command=self.zoom_fit).pack(side="left", padx=2)
        self.lbl_zoom = tk.Label(zoombar, text="100%", font=("Sans", 9),
                                 fg="#999")
        self.lbl_zoom.pack(side="left", padx=6)

        self.lbl_preview = tk.Label(right, bg="#1e1e1e", fg="#888",
                                    text="Load a folder to begin.")
        self.lbl_preview.pack(fill="both", expand=True)

    def _build_bottom(self):
        bot = self._frame(self.root)
        bot.pack(fill="x", padx=8, pady=(4, 8))
        bot.columnconfigure(1, weight=1)

        self._label(bot, "Suggested:").grid(
            row=0, column=0, sticky="w", padx=6, pady=(6, 2))

        self.var_sug = tk.StringVar()
        if USE_CTK:
            self.ent_sug = ctk.CTkEntry(bot, width=760, font=("Mono", 14),
                                        textvariable=self.var_sug)
        else:
            self.ent_sug = tk.Entry(bot, font=("Mono", 12),
                                    textvariable=self.var_sug)
        self.ent_sug.grid(row=0, column=1, sticky="ew", padx=6, pady=2)

        self.lbl_progress = tk.Label(bot, text="", anchor="e")
        self.lbl_progress.grid(row=0, column=2, padx=10)

        btns = tk.Frame(bot)
        btns.grid(row=1, column=1, sticky="w", pady=4)
        self._button(btns, "Approve  (Enter)", self.approve, 140).pack(
            side="left", padx=(0, 6))
        self._button(btns, "Skip  (Tab)", self.next_file, 90).pack(
            side="left", padx=6)
        self._button(btns, "Delete  (Del)", self.delete_file, 100).pack(
            side="left", padx=6)

        self.lbl_help = tk.Label(bot, text=(
            "Single: Enter approves.  Multi (Shift/Ctrl-click): "
            "Enter renames as BaseName - 01, 02, …  Del deletes."),
            font=("Sans", 9), fg="#999")
        self.lbl_help.grid(row=2, column=1, sticky="w", pady=(2, 4))

    def _bind_keys(self):
        self.root.bind("<Return>", lambda e: self.approve())
        self.root.bind("<Tab>", lambda e: self.next_file())
        self.root.bind("<Delete>", self._delete_key)

    # ----------------------------------------------------------- display

    def _refresh_list_display(self, keep_selection_path=None):
        self.lst.delete(0, tk.END)
        for p in self.targets:
            label = os.path.basename(p)
            if p in self.dup_paths:
                label += "   [DUP]"
            self.lst.insert(tk.END, label)
        if keep_selection_path is not None:
            for i, p in enumerate(self.targets):
                if p == keep_selection_path:
                    self.lst.selection_clear(0, tk.END)
                    self.lst.selection_set(i)
                    self.lst.see(i)
                    return

    def _display_preview(self):
        if not getattr(self, "_preview_png", None):
            self.lbl_preview.configure(image="",
                                       text="(no preview available)")
            self.preview_img = None
            return
        try:
            img = tk.PhotoImage(file=self._preview_png)
            p = self.zoom_pct
            if p == 0:      # Fit
                w = img.width()
                if w > 520:
                    sub = max(w // 520, 1)
                    img = img.subsample(sub, sub)
                self.lbl_zoom.configure(text="fit")
            elif p <= 25:
                img = img.subsample(4, 4)
                self.lbl_zoom.configure(text="25%")
            elif p <= 50:
                img = img.subsample(2, 2)
                self.lbl_zoom.configure(text="50%")
            elif p == 150:
                img = img.zoom(3, 3).subsample(2, 2)
                self.lbl_zoom.configure(text="150%")
            elif p == 200:
                img = img.zoom(2, 2)
                self.lbl_zoom.configure(text="200%")
            elif p == 400:
                img = img.zoom(4, 4)
                self.lbl_zoom.configure(text="400%")
            else:
                self.lbl_zoom.configure(text="100%")
            self.preview_img = img
            self.lbl_preview.configure(image=img, text="")
        except Exception as e:
            self.lbl_preview.configure(image="",
                text=f"(preview failed: {e})")

    # -------------------------------------------------------- duplicates

    def _hash_targets(self):
        """Background: md5 every loaded file; flag identical groups."""
        groups = {}
        for p in list(self.targets):
            try:
                h = hashlib.md5()
                with open(p, "rb") as f:
                    for chunk in iter(lambda: f.read(1 << 20), b""):
                        h.update(chunk)
                groups.setdefault(h.hexdigest(), []).append(p)
            except Exception:
                continue
        dups = set()
        for members in groups.values():
            if len(members) > 1:
                dups.update(members)
        self.root.after(0, self._apply_dups, dups)

    def _apply_dups(self, dups):
        self.dup_paths = dups
        self._refresh_list_display(
            keep_selection_path=getattr(self, "current_path", None))
        if dups:
            self.lbl_status.configure(
                text=f"{len(dups)} files share identical content "
                     f"(marked [DUP])")

    def _refresh_list_selection_only(self):
        pass  # reserved

    # ------------------------------------------------------------ events

    def browse(self):
        d = filedialog.askdirectory(initialdir=self.var_dir.get())
        if d:
            self.var_dir.set(d)

    def clear_history(self):
        self.settings["folders"] = []
        save_settings(self.settings)
        if USE_CTK:
            self.cmb_dir.configure(values=[])
        else:
            self.cmb_dir["values"] = []
        messagebox.showinfo("MFDR", "Folder history cleared.")

    def load_folder(self):
        folder = self.var_dir.get()
        if not os.path.isdir(folder):
            messagebox.showerror("MFDR", f"Not a folder:\n{folder}")
            return
        self.settings["last_folder"] = folder
        folders = self.settings.get("folders", [])
        if folder in folders:
            folders.remove(folder)
        folders.insert(0, folder)
        self.settings["folders"] = folders[:20]
        self.settings["recursive"] = self.var_recursive.get()
        self.settings["dry_run"] = self.var_dry_run.get()
        save_settings(self.settings)
        if USE_CTK:
            self.cmb_dir.configure(values=folders[:20])
        else:
            self.cmb_dir["values"] = folders[:20]

        self.targets = mfr.collect_targets(folder, self.var_recursive.get())
        self.dup_paths = set()
        self._refresh_list_display()
        self.lbl_progress.configure(
            text=f"{len(self.targets)} files" if self.targets
                 else "none found")
        self.lbl_preview.configure(text="Select a file to preview.", image="")
        self.var_sug.set("")
        if self.targets:
            self.lst.selection_set(0)
            self.on_select()
        threading.Thread(target=self._hash_targets, daemon=True).start()

    def on_select(self, event=None):
        sel = self.lst.curselection()
        if not sel:
            return
        self.current_path = self.targets[sel[0]]
        n_sel = len(sel)
        if n_sel > 1:
            names = [os.path.basename(self.targets[i]) for i in sel]
            self.lbl_status.configure(
                text=f"{n_sel} files selected — preview shows first:\n"
                     f"  {names[0]}  …  {names[-1]}")
        else:
            self.lbl_status.configure(
                text=f"Analyzing {os.path.basename(self.current_path)} …")
        self.lbl_preview.configure(
            text="rendering…" if n_sel == 1
                 else f"rendering first of {n_sel}…", image="")
        self.preview_img = None
        threading.Thread(target=self._bg_work, daemon=True).start()

    def _bg_work(self):
        path = self.current_path
        if not os.path.exists(path):
            self.root.after(0, self._vanished, path)
            return
        try:
            proposed, meta = mfr.analyze_any(path)
        except Exception as e:
            import traceback
            traceback.print_exc()
            proposed, meta = None, {"method": f"error: {e}"}
        png_path = render_preview_png(path)
        self.root.after(0, self._show_result, path, proposed, meta, png_path)

    def _vanished(self, path):
        self.lbl_status.configure(
            text=f"File vanished from disk:\n  {path}")
        self.lbl_preview.configure(image="", text="(file gone)")
        self.var_sug.set("")

    def _show_result(self, path, proposed, meta, png_path):
        if path != getattr(self, "current_path", None):
            return
        sel = self.lst.curselection()
        n_sel = len(sel)
        if n_sel > 1:
            base = os.path.basename(path)
            if proposed:
                base = os.path.splitext(os.path.basename(proposed))[0]
            self.var_sug.set(base)
            method = meta.get("method", "?") if isinstance(meta, dict) else "?"
            self.lbl_status.configure(
                text=f"{n_sel} files selected — base from first ({method})\n"
                     f"Will rename as: {base} - 01.ext, - 02.ext, …")
        else:
            method = meta.get("method", "?") if isinstance(meta, dict) else "?"
            pages = meta.get("pages", "") if isinstance(meta, dict) else ""
            if proposed:
                self.var_sug.set(os.path.basename(proposed))
                self.lbl_status.configure(
                    text=f"method: {method}   pages: {pages}")
            else:
                self.var_sug.set(os.path.basename(path))
                self.lbl_status.configure(
                    text="(extraction failed — edit below)")
        self._preview_png = png_path if (png_path and
                                         os.path.exists(png_path)) else None
        self._display_preview()

    # ------------------------------------------------------------- zoom

    def zoom_in(self):
        order = [25, 50, 100, 150, 200, 400]
        i = (order.index(self.zoom_pct) + 1 if self.zoom_pct in order
             else 2)
        self.zoom_pct = order[min(i, len(order) - 1)]
        self._display_preview()

    def zoom_out(self):
        order = [25, 50, 100, 150, 200, 400]
        i = (order.index(self.zoom_pct) - 1 if self.zoom_pct in order
             else 2)
        self.zoom_pct = order[max(i, 0)]
        self._display_preview()

    def zoom_fit(self):
        if self._preview_png:
            try:
                img = tk.PhotoImage(file=self._preview_png)
                w = img.width()
                if w > 520:
                    sub = max(w // 520, 1)
                    img = img.subsample(sub, sub)
                self.preview_img = img
                self.lbl_preview.configure(image=img, text="")
                self.lbl_zoom.configure(text="fit")
                return
            except Exception:
                pass
        self._display_preview()

    # ------------------------------------------------------------ rename

    def approve(self):
        sel = self.lst.curselection()
        if not sel:
            return
        if len(sel) == 1:
            self._approve_single(sel[0])
        else:
            self._approve_multi(sel)

    def _approve_single(self, idx):
        path = self.targets[idx]
        new_name = self.var_sug.get().strip()
        if not new_name:
            return
        ext = os.path.splitext(path)[1]
        if not new_name.lower().endswith(ext.lower()):
            new_name += ext
        dest = os.path.join(os.path.dirname(path), mfr.clean(new_name))
        base, e = os.path.splitext(dest)
        n = 1
        while os.path.exists(dest):
            dest = f"{base} ({n}){e}"
            n += 1
        if not self.var_dry_run.get():
            try:
                os.rename(path, dest)
            except OSError as err:
                messagebox.showerror("MFDR", f"Rename failed:\n{err}")
                return
        self.approved_log.append({"from": path, "to": dest})
        self.targets[idx] = dest      # keep position, new name
        self.dup_paths.discard(path)
        self._refresh_list_display(dest)
        self.lbl_progress.configure(
            text=f"processed: {len(self.approved_log)}   "
                 f"remaining: {len(self.targets)}")

    def _approve_multi(self, sel_indices):
        base_name = self.var_sug.get().strip()
        if not base_name:
            messagebox.showwarning("MFDR", "Enter a base name first.")
            return
        pad = max(2, len(str(len(sel_indices))))
        renamed = 0
        for page_num, idx in enumerate(sorted(sel_indices), 1):
            path = self.targets[idx]
            ext = os.path.splitext(path)[1]
            new_name = f"{mfr.clean(base_name)} - {str(page_num).zfill(pad)}{ext}"
            dest = os.path.join(os.path.dirname(path), new_name)
            b, e = os.path.splitext(dest)
            n = 1
            while os.path.exists(dest):
                dest = f"{b} ({n}){e}"
                n += 1
            if not self.var_dry_run.get():
                try:
                    os.rename(path, dest)
                except OSError as err:
                    messagebox.showerror("MFDR",
                                         f"Rename failed on {path}:\n{err}")
                    break
            self.approved_log.append({"from": path, "to": dest})
        # Reload list from disk to reflect the batch
        folder = self.var_dir.get()
        self.targets = mfr.collect_targets(folder, self.var_recursive.get())
        self.dup_paths = set()
        self._refresh_list_display()
        self.lbl_progress.configure(
            text=f"renamed: {len(self.approved_log)}   "
                 f"remaining: {len(self.targets)}")

    def next_file(self):
        if not self.targets:
            self.lbl_status.configure(
                text=f"Done. {len(self.approved_log)} processed.")
            self.lbl_preview.configure(image="", text="Done.")
            self.var_sug.set("")
            return
        sel = self.lst.curselection()
        idx = sel[0] if sel else 0
        idx = min(idx, len(self.targets) - 1)
        self.lst.selection_clear(0, tk.END)
        self.lst.selection_set(idx)
        self.lst.see(idx)
        self.on_select()

    def delete_file(self):
        sel = self.lst.curselection()
        if not sel:
            return
        if len(sel) > 1:
            msg = (f"Delete {len(sel)} selected files permanently?\n\n"
                   "This cannot be undone.")
        else:
            msg = ("Delete this file permanently?\n\n  "
                   f"{os.path.basename(self.targets[sel[0]])}\n\n"
                   "This cannot be undone.")
        if not messagebox.askyesno("MFDR", msg):
            return
        for idx in sorted(sel, reverse=True):
            path = self.targets[idx]
            try:
                if not self.var_dry_run.get():
                    os.remove(path)
                self.approved_log.append({"from": path, "to": "[DELETED]"})
                self.lst.delete(idx)
                self.targets.pop(idx)
            except OSError as err:
                messagebox.showerror("MFDR", f"Delete failed:\n{err}")
                break
        self.lbl_progress.configure(
            text=f"processed: {len(self.approved_log)}   "
                 f"remaining: {len(self.targets)}")
        self.next_file()

    def _delete_key(self, event=None):
        w = self.root.focus_get()
        if w is not None and "entry" in w.winfo_class().lower():
            return  # typing in a text field — Delete edits text
        self.delete_file()

    def write_log(self):
        if not self.approved_log:
            messagebox.showinfo("MFDR", "Nothing approved yet.")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        lp = os.path.join(self.var_dir.get(), f"rename_log_{ts}.json")
        with open(lp, "w") as f:
            json.dump(self.approved_log, f, indent=2)
        messagebox.showinfo(
            "MFDR", f"Log written:\n{lp}\n\nUndo:\n"
                    f"python3 mf_renamer.py --undo {lp}")

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
