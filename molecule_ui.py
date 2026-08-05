"""Drag-and-drop UI for the molecule filter.

Drag a Molecule Explorer export (.xls/.xlsx) onto the window, pick the Priority
and Competition you want, and it writes a new Excel workbook containing:
  * "All Molecules"           - the full source table
  * "<Priority> <Competition>" - just the matching molecules, as a titled table

The extraction logic lives in high_priority_monopoly.py (unchanged); this file
is only the window.
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

from high_priority_monopoly import load_molecules, select_molecules, write_workbook  # noqa: E402

# theme
BG = "#101c2b"
PANEL = "#17273b"
ACCENT = "#2f7fd6"
ACCENT_2 = "#41ab5d"
TEXT = "#e8f0f9"
MUTED = "#93a9c2"
DROP_IDLE = "#1d3552"
DROP_HOVER = "#245079"

PRIORITIES = ["High", "Medium", "Low"]
COMPETITIONS = ["Monopoly", "Duopoly", "Low", "Medium", "High"]

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


class App:
    def __init__(self, root):
        self.root = root
        root.title("Molecule Filter")
        root.configure(bg=BG)
        root.geometry("560x520")
        root.minsize(480, 460)

        self._style()
        self._build()

    def _style(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure("TLabel", background=BG, foreground=TEXT, font=("Helvetica", 12))
        s.configure("Head.TLabel", background=BG, foreground=TEXT, font=("Helvetica", 22, "bold"))
        s.configure("Sub.TLabel", background=BG, foreground=MUTED, font=("Helvetica", 11))
        s.configure("TCombobox", fieldbackground="white", font=("Helvetica", 11))

    def _build(self):
        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True, padx=20, pady=18)

        ttk.Label(outer, text="Molecule Filter", style="Head.TLabel").pack(anchor="w")
        ttk.Label(outer, text="Drag your Molecule Explorer file below to build the filtered sheet.",
                  style="Sub.TLabel", wraplength=500, justify="left").pack(anchor="w", pady=(2, 14))

        # filter selectors
        row = tk.Frame(outer, bg=BG)
        row.pack(fill="x", pady=(0, 12))
        ttk.Label(row, text="Priority:").pack(side="left")
        self.priority = ttk.Combobox(row, values=PRIORITIES, state="readonly", width=10)
        self.priority.set("High")
        self.priority.pack(side="left", padx=(6, 18))
        ttk.Label(row, text="Competition:").pack(side="left")
        self.competition = ttk.Combobox(row, values=COMPETITIONS, state="readonly", width=12)
        self.competition.set("Monopoly")
        self.competition.pack(side="left", padx=(6, 0))

        # drop zone
        self.drop = tk.Label(
            outer, text="⬇  Drop the Excel file here\n(or click to browse)",
            bg=DROP_IDLE, fg=TEXT, font=("Helvetica", 14), height=6, cursor="hand2", bd=2,
        )
        self.drop.pack(fill="both", expand=True, pady=(0, 12))
        self.drop.bind("<Button-1>", lambda e: self.browse())
        if _DND:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self.on_drop)
            self.drop.dnd_bind("<<DragEnter>>", lambda e: self.drop.configure(bg=DROP_HOVER))
            self.drop.dnd_bind("<<DragLeave>>", lambda e: self.drop.configure(bg=DROP_IDLE))

        self.progress = ttk.Progressbar(outer, mode="indeterminate")
        self.status = ttk.Label(outer, text="Waiting for a file…", style="Sub.TLabel", wraplength=500, justify="left")
        self.status.pack(anchor="w")

        self.reveal_path = None
        self.reveal_btn = tk.Button(
            outer, text=REVEAL_LABEL, command=self.reveal, bg=ACCENT_2, fg="white",
            relief="flat", bd=0, font=("Helvetica", 12, "bold"), cursor="hand2",
        )

    # --- intake ---
    def browse(self):
        path = filedialog.askopenfilename(
            title="Choose a Molecule Explorer file",
            filetypes=[("Excel / molecule export", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.process(path)

    def on_drop(self, event):
        self.drop.configure(bg=DROP_IDLE)
        paths = self.root.tk.splitlist(event.data)
        files = [p for p in paths if p.lower().endswith((".xls", ".xlsx"))]
        if not files:
            messagebox.showwarning("Not an Excel file", "Please drop a .xls or .xlsx file.")
            return
        self.process(files[0])

    def process(self, path):
        self.reveal_btn.pack_forget()
        self.status.configure(text=f"Processing {os.path.basename(path)}…")
        self.progress.pack(fill="x", pady=(4, 8))
        self.progress.start(12)
        priority = self.priority.get() or "High"
        competition = self.competition.get() or "Monopoly"
        threading.Thread(target=self._run, args=(path, priority, competition), daemon=True).start()

    def _run(self, path, priority, competition):
        try:
            df = load_molecules(path)
            selected = select_molecules(df, priority, competition)
            stem = os.path.splitext(os.path.basename(path))[0]
            out = os.path.join(os.path.dirname(path), f"{stem} - {priority} {competition}.xlsx")
            sheet = f"{priority} {competition}"[:31]
            write_workbook(df, selected, out, priority, competition, sheet_name=sheet)
            self.root.after(0, lambda: self._done(out, len(selected), len(df), priority, competition))
        except Exception as exc:
            tb = traceback.format_exc()
            self.root.after(0, lambda: self._error(exc, tb))

    def _done(self, out, n, total, priority, competition):
        self.progress.stop()
        self.progress.pack_forget()
        self.reveal_path = out
        self.status.configure(
            text=f"✓ Found {n} {priority} + {competition} molecules (of {total}).\n"
                 f"Saved to: {out}\nSheets: 'All Molecules' and '{priority} {competition}'."
        )
        self.reveal_btn.pack(anchor="w", pady=(10, 0))

    def _error(self, exc, tb):
        self.progress.stop()
        self.progress.pack_forget()
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
