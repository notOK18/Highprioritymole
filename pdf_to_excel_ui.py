"""Drag-and-drop UI to convert PDFs into one Excel workbook.

Drop one or more PDFs (or click to browse). Each PDF becomes a sheet: its tables
go into columns, or its text if it has no tables. Click Convert to Excel to write
a single workbook. The logic lives in pdf_to_excel.py.
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

from pdf_to_excel import convert  # noqa: E402

BG = "#101c2b"
ACCENT = "#2f7fd6"
ACCENT_2 = "#41ab5d"
DANGER = "#b3455a"
TEXT = "#e8f0f9"
MUTED = "#93a9c2"
DROP_IDLE = "#1d3552"
DROP_HOVER = "#245079"
PANEL = "#17273b"

REVEAL_LABEL = {"darwin": "Show in Finder", "win32": "Show in Folder"}.get(sys.platform, "Open Folder")


def reveal_in_file_manager(path):
    """Open the OS file manager with `path` selected (macOS/Windows/Linux)."""
    if not (path and os.path.exists(path)):
        return
    if sys.platform == "darwin":
        subprocess.run(["open", "-R", path])
    elif sys.platform == "win32":
        subprocess.run(f'explorer /select,"{os.path.normpath(path)}"')
    else:
        subprocess.run(["xdg-open", os.path.dirname(path)])


class App:
    def __init__(self, root):
        self.root = root
        self.paths = []  # accumulated PDF paths, in order added
        root.title("PDF to Excel")
        root.configure(bg=BG)
        root.geometry("580x640")
        root.minsize(520, 560)
        self._style()

        outer = tk.Frame(root, bg=BG)
        outer.pack(fill="both", expand=True, padx=20, pady=18)

        tk.Label(outer, text="PDF to Excel", bg=BG, fg=TEXT,
                 font=("Helvetica", 22, "bold")).pack(anchor="w")
        tk.Label(outer, text="Drop PDFs below (add as many as you like). Each becomes a sheet in one workbook.",
                 bg=BG, fg=MUTED, font=("Helvetica", 11), wraplength=520, justify="left").pack(anchor="w", pady=(2, 14))

        self.box = tk.Label(outer, text="⬇  Drop PDF files here  ·  or click to browse",
                            bg=DROP_IDLE, fg=TEXT, font=("Helvetica", 12), height=3, cursor="hand2", bd=2)
        self.box.pack(fill="x")
        self.box.bind("<Button-1>", lambda e: self.browse())
        if _DND:
            self.box.drop_target_register(DND_FILES)
            self.box.dnd_bind("<<Drop>>", self.on_drop)
            self.box.dnd_bind("<<DragEnter>>", lambda e: self.box.configure(bg=DROP_HOVER))
            self.box.dnd_bind("<<DragLeave>>", lambda e: self.box.configure(bg=DROP_IDLE))

        # list of added files
        list_wrap = tk.Frame(outer, bg=PANEL, bd=0)
        list_wrap.pack(fill="both", expand=True, pady=(10, 8))
        self.listbox = tk.Listbox(list_wrap, bg=PANEL, fg=TEXT, height=7, bd=0,
                                  highlightthickness=0, selectbackground=ACCENT, activestyle="none",
                                  font=("Helvetica", 11))
        self.listbox.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        sb = ttk.Scrollbar(list_wrap, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=sb.set)

        btn_row = tk.Frame(outer, bg=BG)
        btn_row.pack(fill="x")
        tk.Button(btn_row, text="Remove selected", command=self.remove_selected, bg=PANEL, fg=TEXT,
                  relief="flat", bd=0, font=("Helvetica", 11), cursor="hand2").pack(side="left")
        tk.Button(btn_row, text="Clear all", command=self.clear_all, bg=PANEL, fg=TEXT,
                  relief="flat", bd=0, font=("Helvetica", 11), cursor="hand2").pack(side="left", padx=(8, 0))

        self.button = tk.Button(outer, text="Convert to Excel", command=self.run,
                                bg=ACCENT, fg="white", activebackground=ACCENT_2, activeforeground="white",
                                disabledforeground="#7f97b3", font=("Helvetica", 15, "bold"),
                                relief="flat", bd=0, height=2, state="disabled", cursor="hand2")
        self.button.pack(fill="x", pady=(10, 8))

        self.progress = ttk.Progressbar(outer, mode="indeterminate")
        self.status = tk.Label(outer, text="No PDFs added yet.", bg=BG, fg=MUTED,
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

    # ---- file list management ----
    def add_paths(self, paths):
        added = 0
        for p in paths:
            if p.lower().endswith(".pdf") and p not in self.paths:
                self.paths.append(p)
                self.listbox.insert("end", "  " + os.path.basename(p))
                added += 1
        self.refresh(added)

    def browse(self):
        paths = filedialog.askopenfilenames(
            title="Choose PDF files", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if paths:
            self.add_paths(list(paths))

    def on_drop(self, event):
        self.box.configure(bg=DROP_IDLE)
        paths = self.root.tk.splitlist(event.data)
        pdfs = [p for p in paths if p.lower().endswith(".pdf")]
        if not pdfs:
            messagebox.showwarning("No PDFs", "Please drop .pdf files.")
            return
        self.add_paths(pdfs)

    def remove_selected(self):
        for i in sorted(self.listbox.curselection(), reverse=True):
            self.listbox.delete(i)
            del self.paths[i]
        self.refresh()

    def clear_all(self):
        self.listbox.delete(0, "end")
        self.paths.clear()
        self.refresh()

    def refresh(self, just_added=None):
        n = len(self.paths)
        self.button.configure(state="normal" if n else "disabled")
        if n:
            extra = f"  (+{just_added} added)" if just_added else ""
            self.status.configure(text=f"{n} PDF{'s' if n != 1 else ''} ready.{extra} Click Convert to Excel.")
        else:
            self.status.configure(text="No PDFs added yet.")

    # ---- conversion ----
    def _output_path(self):
        first_dir = os.path.dirname(self.paths[0])
        if len(self.paths) == 1:
            name = os.path.splitext(os.path.basename(self.paths[0]))[0] + ".xlsx"
        else:
            name = f"PDF export ({len(self.paths)} files).xlsx"
        return os.path.join(first_dir, name)

    def run(self):
        self.reveal_btn.pack_forget()
        self.button.configure(state="disabled")
        self.status.configure(text="Converting… (scanned PDFs can take a moment)")
        self.progress.pack(fill="x", pady=(4, 8))
        self.progress.start(12)
        threading.Thread(target=self._work, args=(list(self.paths), self._output_path()), daemon=True).start()

    def _work(self, paths, out):
        try:
            _, sheets = convert(paths, out)
            self.root.after(0, lambda: self._done(out, sheets))
        except Exception as exc:
            tb = traceback.format_exc()
            self.root.after(0, lambda: self._error(exc, tb))

    def _done(self, out, sheets):
        self.progress.stop()
        self.progress.pack_forget()
        self.button.configure(state="normal")
        self.reveal_path = out
        self.status.configure(text=f"✓ Wrote {sheets} sheet{'s' if sheets != 1 else ''}.\nSaved to: {out}")
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
