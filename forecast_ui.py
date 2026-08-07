"""Drag-and-drop UI to enrich supplier-common molecules with Forecast data.

Drop the "Supplier common molecules" workbook in the first box and point the
second box at your Forecast folder, then click Build Forecast Data. For every
common molecule it pulls MIDAS sales (2023/2024/2025), hospital Lebanon
consumption (40% + total), and tender min/max quantities, and writes a workbook
with a Summary sheet plus a detail sheet per source. Logic: forecast_enrich.py.
"""

import os
import subprocess
import sys
import threading
import traceback

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND = True
except Exception:
    _DND = False

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from forecast_enrich import enrich  # noqa: E402
from compare_molecules import list_sheets  # noqa: E402

BG = "#101c2b"
ACCENT = "#2f7fd6"
ACCENT_2 = "#41ab5d"
TEXT = "#e8f0f9"
MUTED = "#93a9c2"
DROP_IDLE = "#1d3552"
DROP_HOVER = "#245079"
DROP_SET = "#1f4a36"

# "Show in Finder" is macOS wording; on Windows/Linux the file manager differs.
REVEAL_LABEL = {"darwin": "Show in Finder", "win32": "Show in Folder"}.get(sys.platform, "Open Folder")


def reveal_in_file_manager(path):
    """Open the OS file manager with `path` selected (macOS/Windows/Linux)."""
    if not (path and os.path.exists(path)):
        return
    if sys.platform == "darwin":
        subprocess.run(["open", "-R", path])
    elif sys.platform == "win32":
        subprocess.run(["explorer", "/select,", os.path.normpath(path)])
    else:
        subprocess.run(["xdg-open", os.path.dirname(path)])


class FileZone:
    """A drop/browse box for a single file, with an optional sheet chooser."""

    def __init__(self, app, parent, title, hint, default_sheet=None):
        self.app = app
        self.default_sheet = default_sheet
        self.path = None

        wrap = tk.Frame(parent, bg=BG)
        wrap.pack(fill="x", pady=(0, 10))
        tk.Label(wrap, text=title, bg=BG, fg=TEXT, font=("Helvetica", 12, "bold")).pack(anchor="w")
        tk.Label(wrap, text=hint, bg=BG, fg=MUTED, font=("Helvetica", 10)).pack(anchor="w", pady=(0, 4))

        self.box = tk.Label(wrap, text="⬇  Drop file here  ·  or click to browse",
                            bg=DROP_IDLE, fg=TEXT, font=("Helvetica", 12), height=2, cursor="hand2", bd=2)
        self.box.pack(fill="x")
        self.box.bind("<Button-1>", lambda e: self.browse())
        if _DND:
            self.box.drop_target_register(DND_FILES)
            self.box.dnd_bind("<<Drop>>", self.on_drop)
            self.box.dnd_bind("<<DragEnter>>", lambda e: self.box.configure(bg=DROP_HOVER))
            self.box.dnd_bind("<<DragLeave>>", lambda e: self.box.configure(bg=self._idle_bg()))

        self.sheet_row = tk.Frame(wrap, bg=BG)
        tk.Label(self.sheet_row, text="Sheet:", bg=BG, fg=MUTED, font=("Helvetica", 11)).pack(side="left")
        self.sheet_cb = ttk.Combobox(self.sheet_row, state="readonly", width=32)
        self.sheet_cb.pack(side="left", padx=(6, 0))

    def _idle_bg(self):
        return DROP_SET if self.path else DROP_IDLE

    def browse(self):
        path = filedialog.askopenfilename(
            title="Choose the Supplier common molecules file",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")])
        if path:
            self.set_path(path)

    def on_drop(self, event):
        paths = self.app.root.tk.splitlist(event.data)
        files = [p for p in paths if p.lower().endswith((".xls", ".xlsx"))]
        if not files:
            messagebox.showwarning("Not an Excel file", "Please drop a .xls or .xlsx file.")
            self.box.configure(bg=self._idle_bg())
            return
        self.set_path(files[0])

    def set_path(self, path):
        self.path = path
        self.box.configure(text="📄  " + os.path.basename(path), bg=DROP_SET)
        sheets = list_sheets(path)
        if sheets:
            self.sheet_cb["values"] = sheets
            default = next((s for s in sheets if s.lower() == (self.default_sheet or "").lower()), sheets[0])
            self.sheet_cb.set(default)
            self.sheet_row.pack(fill="x", pady=(6, 0))
        else:
            self.sheet_cb.set("")
            self.sheet_row.pack_forget()
        self.app.autofill_forecast(path)
        self.app.refresh()

    def sheet(self):
        return self.sheet_cb.get() or None


class FolderZone:
    """A drop/browse box for a folder."""

    def __init__(self, app, parent, title, hint):
        self.app = app
        self.path = None

        wrap = tk.Frame(parent, bg=BG)
        wrap.pack(fill="x", pady=(0, 10))
        tk.Label(wrap, text=title, bg=BG, fg=TEXT, font=("Helvetica", 12, "bold")).pack(anchor="w")
        tk.Label(wrap, text=hint, bg=BG, fg=MUTED, font=("Helvetica", 10)).pack(anchor="w", pady=(0, 4))

        self.box = tk.Label(wrap, text="⬇  Drop folder here  ·  or click to browse",
                            bg=DROP_IDLE, fg=TEXT, font=("Helvetica", 12), height=2, cursor="hand2", bd=2)
        self.box.pack(fill="x")
        self.box.bind("<Button-1>", lambda e: self.browse())
        if _DND:
            self.box.drop_target_register(DND_FILES)
            self.box.dnd_bind("<<Drop>>", self.on_drop)
            self.box.dnd_bind("<<DragEnter>>", lambda e: self.box.configure(bg=DROP_HOVER))
            self.box.dnd_bind("<<DragLeave>>", lambda e: self.box.configure(bg=self._idle_bg()))

    def _idle_bg(self):
        return DROP_SET if self.path else DROP_IDLE

    def browse(self):
        path = filedialog.askdirectory(title="Choose the Forecast folder")
        if path:
            self.set_path(path)

    def on_drop(self, event):
        paths = self.app.root.tk.splitlist(event.data)
        folders = [p for p in paths if os.path.isdir(p)]
        if not folders:
            messagebox.showwarning("Not a folder", "Please drop a folder, not a file.")
            self.box.configure(bg=self._idle_bg())
            return
        self.set_path(folders[0])

    def set_path(self, path):
        self.path = path
        self.box.configure(text="📁  " + os.path.basename(path.rstrip("/\\")) or path, bg=DROP_SET)
        self.app.refresh()


class App:
    def __init__(self, root):
        self.root = root
        root.title("Forecast Data Builder")
        root.configure(bg=BG)
        root.geometry("580x620")
        root.minsize(520, 560)
        self._style()

        outer = tk.Frame(root, bg=BG)
        outer.pack(fill="both", expand=True, padx=20, pady=18)

        tk.Label(outer, text="Forecast Data Builder", bg=BG, fg=TEXT,
                 font=("Helvetica", 22, "bold")).pack(anchor="w")
        tk.Label(outer, text="Enrich the supplier-common molecules with MIDAS sales, hospital "
                             "consumption, and tender quantities.",
                 bg=BG, fg=MUTED, font=("Helvetica", 11), wraplength=520, justify="left").pack(anchor="w", pady=(2, 14))

        self.common = FileZone(self, outer, "1 · Supplier common molecules",
                               "The workbook from the Supplier Common Molecules app.", default_sheet="Common")
        self.forecast = FolderZone(self, outer, "2 · Forecast folder",
                                   "The folder holding the MIDAS, hospital, and tender files.")

        self.button = tk.Button(outer, text="Build Forecast Data", command=self.run,
                                bg=ACCENT, fg="white", activebackground=ACCENT_2, activeforeground="white",
                                disabledforeground="#7f97b3", font=("Helvetica", 15, "bold"),
                                relief="flat", bd=0, height=2, state="disabled", cursor="hand2")
        self.button.pack(fill="x", pady=(6, 8))

        self.progress = ttk.Progressbar(outer, mode="indeterminate")
        self.status = tk.Label(outer, text="Waiting for the file and the Forecast folder…", bg=BG, fg=MUTED,
                               font=("Helvetica", 11), wraplength=520, justify="left")
        self.status.pack(anchor="w")

        self.reveal_path = None
        self.reveal_btn = tk.Button(outer, text=REVEAL_LABEL, command=self.reveal,
                                    bg=ACCENT_2, fg="white", relief="flat", bd=0,
                                    font=("Helvetica", 12, "bold"), cursor="hand2")

    def _style(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass

    def autofill_forecast(self, common_path):
        """If a 'Forecast' folder sits near the chosen file, prefill it."""
        if self.forecast.path:
            return
        base = os.path.dirname(common_path)
        for cand in (os.path.join(base, "Forecast"), os.path.join(os.path.dirname(base), "Forecast")):
            if os.path.isdir(cand):
                self.forecast.set_path(cand)
                return

    def refresh(self):
        ready = bool(self.common.path and self.forecast.path)
        self.button.configure(state="normal" if ready else "disabled")
        if ready:
            self.status.configure(text="Ready. Click Build Forecast Data.")

    def run(self):
        self.reveal_btn.pack_forget()
        self.button.configure(state="disabled")
        self.status.configure(text="Reading forecast files (the tender PDF can take a moment)…")
        self.progress.pack(fill="x", pady=(4, 8))
        self.progress.start(12)
        args = (self.common.path, self.common.sheet(), self.forecast.path)
        threading.Thread(target=self._work, args=args, daemon=True).start()

    def _work(self, common_path, common_sheet, forecast_dir):
        try:
            stem = os.path.splitext(os.path.basename(common_path))[0]
            out = os.path.join(os.path.dirname(common_path), f"Forecast data - {stem}.xlsx")
            _, matched = enrich(common_path, forecast_dir, out, common_sheet=common_sheet)
            self.root.after(0, lambda: self._done(out, matched))
        except Exception as exc:
            tb = traceback.format_exc()
            self.root.after(0, lambda: self._error(exc, tb))

    def _done(self, out, matched):
        self.progress.stop()
        self.progress.pack_forget()
        self.button.configure(state="normal")
        self.reveal_path = out
        self.status.configure(
            text=f"✓ {matched} common molecules had forecast data.\nSaved to: {out}\n"
                 f"Sheets: 'Summary', 'MIDAS detail', 'Hospital detail', 'Tender detail'.")
        self.reveal_btn.pack(anchor="w", pady=(10, 0))

    def _error(self, exc, tb):
        self.progress.stop()
        self.progress.pack_forget()
        self.button.configure(state="normal")
        self.status.configure(text="Something went wrong.")
        messagebox.showerror("Error", f"{exc}\n\n{tb}")

    def reveal(self):
        reveal_in_file_manager(self.reveal_path)


def main():
    root = TkinterDnD.Tk() if _DND else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
