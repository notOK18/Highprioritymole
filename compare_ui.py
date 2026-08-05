"""Drag-and-drop UI to find molecules common to two lists.

Drop a "Potential molecules" file in the top box and a "High Priority +
Monopoly molecules" file in the bottom box, then click Find Common Molecules.
It writes a workbook with Potential / High Priority Monopoly / Common Molecules
sheets. The comparison logic lives in compare_molecules.py (unchanged here).
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

from compare_molecules import compare_files, list_sheets  # noqa: E402

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
        # explorer returns exit code 1 even on success, so don't check it.
        subprocess.run(["explorer", "/select,", os.path.normpath(path)])
    else:
        subprocess.run(["xdg-open", os.path.dirname(path)])


class DropZone:
    def __init__(self, app, parent, title, hint, keyword):
        self.app = app
        self.keyword = keyword
        self.path = None

        wrap = tk.Frame(parent, bg=BG)
        wrap.pack(fill="x", pady=(0, 12))
        tk.Label(wrap, text=title, bg=BG, fg=TEXT, font=("Helvetica", 13, "bold")).pack(anchor="w")
        tk.Label(wrap, text=hint, bg=BG, fg=MUTED, font=("Helvetica", 10)).pack(anchor="w", pady=(0, 4))

        self.box = tk.Label(
            wrap, text="⬇  Drop file here  ·  or click to browse",
            bg=DROP_IDLE, fg=TEXT, font=("Helvetica", 12), height=3, cursor="hand2", bd=2,
        )
        self.box.pack(fill="x")
        self.box.bind("<Button-1>", lambda e: self.browse())
        if _DND:
            self.box.drop_target_register(DND_FILES)
            self.box.dnd_bind("<<Drop>>", self.on_drop)
            self.box.dnd_bind("<<DragEnter>>", lambda e: self.box.configure(bg=DROP_HOVER))
            self.box.dnd_bind("<<DragLeave>>", lambda e: self.box.configure(bg=self._idle_bg()))

        # sheet chooser (shown only when the file has selectable sheets)
        self.sheet_row = tk.Frame(wrap, bg=BG)
        tk.Label(self.sheet_row, text="Sheet:", bg=BG, fg=MUTED, font=("Helvetica", 11)).pack(side="left")
        self.sheet_cb = ttk.Combobox(self.sheet_row, state="readonly", width=32)
        self.sheet_cb.pack(side="left", padx=(6, 0))

    def _idle_bg(self):
        return DROP_SET if self.path else DROP_IDLE

    def browse(self):
        path = filedialog.askopenfilename(
            title="Choose an Excel file",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
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

        # populate the sheet chooser for real multi-sheet workbooks
        sheets = list_sheets(path)
        if sheets:
            self.sheet_cb["values"] = sheets
            # default to a sheet whose name hints at this list's role
            default = next((s for s in sheets if self.keyword in s.lower()), sheets[0])
            self.sheet_cb.set(default)
            self.sheet_row.pack(fill="x", pady=(6, 0))
        else:
            self.sheet_cb.set("")
            self.sheet_row.pack_forget()  # HTML-.xls export: single table, no choice
        self.app.refresh()

    def sheet(self):
        """The chosen sheet name, or None to auto-detect."""
        value = self.sheet_cb.get()
        return value or None


class App:
    def __init__(self, root):
        self.root = root
        root.title("Common Molecules Finder")
        root.configure(bg=BG)
        root.geometry("560x560")
        root.minsize(500, 520)
        self._style()

        outer = tk.Frame(root, bg=BG)
        outer.pack(fill="both", expand=True, padx=20, pady=18)

        tk.Label(outer, text="Common Molecules Finder", bg=BG, fg=TEXT,
                 font=("Helvetica", 22, "bold")).pack(anchor="w")
        tk.Label(outer, text="Drop both lists, then find the molecules that appear in both.",
                 bg=BG, fg=MUTED, font=("Helvetica", 11)).pack(anchor="w", pady=(2, 16))

        self.potential = DropZone(self, outer, "1 · Potential molecules",
                                  "Your list of potential molecules.", "potential")
        self.hp = DropZone(self, outer, "2 · High Priority + Monopoly molecules",
                           "The high-priority, monopoly-competition list.", "monopoly")

        self.button = tk.Button(
            outer, text="Find Common Molecules", command=self.run,
            bg=ACCENT, fg="white", activebackground=ACCENT_2, activeforeground="white",
            disabledforeground="#7f97b3", font=("Helvetica", 15, "bold"),
            relief="flat", bd=0, height=2, state="disabled", cursor="hand2",
        )
        self.button.pack(fill="x", pady=(6, 8))

        self.progress = ttk.Progressbar(outer, mode="indeterminate")
        self.status = tk.Label(outer, text="Waiting for both files…", bg=BG, fg=MUTED,
                               font=("Helvetica", 11), wraplength=500, justify="left")
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

    def refresh(self):
        ready = bool(self.potential.path and self.hp.path)
        self.button.configure(state="normal" if ready else "disabled")
        if ready:
            self.status.configure(text="Ready. Click Find Common Molecules.")

    def run(self):
        self.reveal_btn.pack_forget()
        self.button.configure(state="disabled")
        self.status.configure(text="Comparing…")
        self.progress.pack(fill="x", pady=(4, 8))
        self.progress.start(12)
        args = (self.potential.path, self.hp.path, self.potential.sheet(), self.hp.sheet())
        threading.Thread(target=self._work, args=args, daemon=True).start()

    def _work(self, potential_path, hp_path, potential_sheet, hp_sheet):
        try:
            out = os.path.join(os.path.dirname(potential_path), "Common Molecules.xlsx")
            _, count = compare_files(potential_path, hp_path, out,
                                     potential_sheet=potential_sheet, hp_sheet=hp_sheet)
            self.root.after(0, lambda: self._done(out, count))
        except Exception as exc:
            tb = traceback.format_exc()
            self.root.after(0, lambda: self._error(exc, tb))

    def _done(self, out, count):
        self.progress.stop()
        self.progress.pack_forget()
        self.button.configure(state="normal")
        self.reveal_path = out
        self.status.configure(
            text=f"✓ {count} molecules are in BOTH lists.\nSaved to: {out}\n"
                 f"Sheets: 'Potential', 'High Priority Monopoly', 'Common Molecules'."
        )
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
