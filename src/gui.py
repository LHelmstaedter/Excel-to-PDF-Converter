from __future__ import annotations

import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

# Allow running this file directly (e.g. via the editor's Play button).
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.pdf_exporter import export_group_pdfs  # noqa: E402

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    _DND_AVAILABLE = True
except Exception:
    DND_FILES = TkinterDnD = None
    _DND_AVAILABLE = False

BaseTk = TkinterDnD.Tk if _DND_AVAILABLE else tk.Tk


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title("Excel → PDFs (Gruppenweise)")
        self.geometry("760x420")
        self.resizable(False, False)

        self.excel_path = ""
        self.out_dir = ""
        self.excel_label = tk.StringVar(value="")
        self.out_label = tk.StringVar(value="")
        self.status = tk.StringVar(value="Bitte Excel-Datei auswählen (Drag & Drop optional).")

        tk.Label(self, text="Excel-Datei einfügen", font=("Segoe UI", 14, "bold")).pack(pady=12)

        drop_text = (
            "Drop-Zone (Drag & Drop aktiv)\n\nZiehe eine .xlsx/.xlsm Datei hier hinein\noder klicke auf 'Datei auswählen…'."
            if _DND_AVAILABLE else
            "Drag & Drop nicht verfügbar\n\nBitte über 'Datei auswählen…' die Excel-Datei laden."
        )
        drop = tk.Label(self, text=drop_text, relief="groove", borderwidth=2, width=72, height=10)
        drop.pack(pady=8)
        if _DND_AVAILABLE:
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind("<<Drop>>", self.on_drop)

        row = tk.Frame(self)
        row.pack(pady=6)
        tk.Button(row, text="Datei auswählen…", command=self.on_browse).pack(side="left", padx=6)
        tk.Button(row, text="Ausgabeordner…", command=self.on_outdir).pack(side="left", padx=6)
        self.run_btn = tk.Button(row, text="PDFs erzeugen", command=self.on_run)
        self.run_btn.pack(side="left", padx=6)

        tk.Label(self, textvariable=self.excel_label, wraplength=720).pack(pady=6)
        tk.Label(self, textvariable=self.out_label, wraplength=720).pack(pady=2)
        tk.Label(self, textvariable=self.status, fg="gray", wraplength=720).pack(pady=6)
        tk.Label(self, text="Speicherstruktur: <Prefix>/<Prefix><Term>/...", fg="gray").pack(side="bottom", pady=10)

    def on_drop(self, event):
        self.set_excel(self.tk.splitlist(event.data)[0])  # handles {paths with spaces}

    def on_browse(self):
        path = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            self.set_excel(path)

    def on_outdir(self):
        d = filedialog.askdirectory(title="Ausgabeordner wählen")
        if d:
            self.out_dir = d
            self.out_label.set(f"Ausgabe: {d}")

    def set_excel(self, path: str):
        p = Path(path)
        if not p.exists():
            messagebox.showerror("Fehler", "Datei nicht gefunden.")
        elif p.suffix.lower() not in (".xlsx", ".xlsm", ".xls"):
            messagebox.showerror("Fehler", "Bitte eine Excel-Datei auswählen (.xlsx/.xlsm/.xls).")
        else:
            self.excel_path = str(p)
            self.excel_label.set(str(p))
            self.status.set("Bereit. Klicke 'PDFs erzeugen'.")

    def on_run(self):
        if not self.excel_path:
            messagebox.showwarning("Hinweis", "Bitte zuerst eine Excel-Datei auswählen oder hineinziehen.")
            return

        excel_path = str(Path(self.excel_path).resolve())  # Excel needs an absolute path
        out_dir = self.out_dir or str(Path(excel_path).parent / "pdf_out")

        self.run_btn.config(state="disabled")
        self.status.set("Erzeuge PDFs…")

        def worker():
            error = None
            try:
                export_group_pdfs(excel_path=excel_path, out_dir=out_dir)
            except Exception as e:
                error = str(e)  # 'e' is unset after the except block, so keep a copy
            self.after(0, lambda: self.done(out_dir, error))

        threading.Thread(target=worker, daemon=True).start()

    def done(self, out_dir: str, error: str | None):
        self.run_btn.config(state="normal")
        if error:
            self.status.set("Fehler.")
            messagebox.showerror("Fehler", error)
        else:
            self.status.set("Fertig.")
            messagebox.showinfo("Fertig", f"PDFs erstellt in:\n{out_dir}")


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
