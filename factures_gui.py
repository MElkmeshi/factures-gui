#!/usr/bin/env python3
"""
Simple desktop GUI for generate_factures.py.

Pick a workbook, set the year and output folder, choose what to generate, and
press "Générer". The underlying generate_factures.py script is run as a
subprocess and its progress is streamed live into the log box.

Cross-platform (macOS / Windows / Linux). Requires Python with tkinter
(bundled with the standard python.org installers) plus whatever
generate_factures.py needs (openpyxl + LibreOffice for PDFs).

Run:  python3 factures_gui.py
"""
import os
import queue
import subprocess
import sys
import threading
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

HERE = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = HERE / "Factures Livreurs.xlsx"


class FacturesGUI:
    def __init__(self, root):
        self.root = root
        root.title("Générateur de factures")
        root.minsize(640, 480)

        self.proc = None
        self.log_queue = queue.Queue()
        self.last_out_dir = None

        self._build_widgets()
        self.root.after(100, self._drain_log)

    # -- layout ---------------------------------------------------------------
    def _build_widgets(self):
        pad = {"padx": 8, "pady": 4}
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        # Workbook
        ttk.Label(frm, text="Classeur (.xlsx) :").grid(row=0, column=0, sticky="w", **pad)
        self.workbook_var = tk.StringVar(
            value=str(DEFAULT_WORKBOOK) if DEFAULT_WORKBOOK.exists() else ""
        )
        ttk.Entry(frm, textvariable=self.workbook_var).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Parcourir…", command=self._pick_workbook).grid(row=0, column=2, **pad)

        # Year
        ttk.Label(frm, text="Année :").grid(row=1, column=0, sticky="w", **pad)
        self.year_var = tk.IntVar(value=2025)
        ttk.Spinbox(frm, from_=2020, to=2099, textvariable=self.year_var, width=8).grid(
            row=1, column=1, sticky="w", **pad
        )

        # Output folder
        ttk.Label(frm, text="Dossier de sortie :").grid(row=2, column=0, sticky="w", **pad)
        self.outdir_var = tk.StringVar(value="")
        ttk.Entry(frm, textvariable=self.outdir_var).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Parcourir…", command=self._pick_outdir).grid(row=2, column=2, **pad)
        ttk.Label(frm, text="(laisser vide = à côté du classeur)", foreground="gray").grid(
            row=3, column=1, sticky="w", padx=8
        )

        # Output toggles
        opts = ttk.LabelFrame(frm, text="Options", padding=8)
        opts.grid(row=4, column=0, columnspan=3, sticky="ew", **pad)
        self.pdf_var = tk.BooleanVar(value=True)
        self.excel_var = tk.BooleanVar(value=True)
        self.keep_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="Générer les PDF (un par livreur)", variable=self.pdf_var).pack(anchor="w")
        ttk.Checkbutton(opts, text="Générer le classeur Excel combiné", variable=self.excel_var).pack(anchor="w")
        ttk.Checkbutton(opts, text="Conserver les .xlsx intermédiaires", variable=self.keep_var).pack(anchor="w")

        # Buttons
        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=3, sticky="ew", **pad)
        self.run_btn = ttk.Button(btns, text="Générer", command=self._run)
        self.run_btn.pack(side="left")
        self.open_btn = ttk.Button(btns, text="Ouvrir le dossier", command=self._open_outdir, state="disabled")
        self.open_btn.pack(side="left", padx=8)

        # Log
        ttk.Label(frm, text="Journal :").grid(row=6, column=0, sticky="w", padx=8)
        self.log = tk.Text(frm, height=14, wrap="word", state="disabled")
        self.log.grid(row=7, column=0, columnspan=3, sticky="nsew", **pad)
        frm.rowconfigure(7, weight=1)
        scroll = ttk.Scrollbar(frm, command=self.log.yview)
        scroll.grid(row=7, column=3, sticky="ns")
        self.log["yscrollcommand"] = scroll.set

    # -- pickers --------------------------------------------------------------
    def _pick_workbook(self):
        path = filedialog.askopenfilename(
            title="Choisir le classeur",
            filetypes=[("Classeurs Excel", "*.xlsx"), ("Tous les fichiers", "*.*")],
        )
        if path:
            self.workbook_var.set(path)

    def _pick_outdir(self):
        path = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if path:
            self.outdir_var.set(path)

    def _open_outdir(self):
        target = self.last_out_dir
        if not target or not Path(target).exists():
            messagebox.showinfo("Info", "Aucun dossier de sortie disponible pour le moment.")
            return
        if sys.platform == "darwin":
            subprocess.run(["open", str(target)])
        elif os.name == "nt":
            os.startfile(str(target))  # noqa: S606
        else:
            subprocess.run(["xdg-open", str(target)])

    # -- run ------------------------------------------------------------------
    def _run(self):
        workbook = self.workbook_var.get().strip()
        if not workbook or not Path(workbook).exists():
            messagebox.showerror("Erreur", "Veuillez choisir un classeur .xlsx valide.")
            return
        if not self.pdf_var.get() and not self.excel_var.get():
            messagebox.showerror("Erreur", "Choisissez au moins un type de sortie (PDF ou Excel).")
            return

        params = {
            "workbook": workbook,
            "year": int(self.year_var.get()),
            "out": self.outdir_var.get().strip() or None,
            "keep_xlsx": self.keep_var.get(),
            "no_pdf": not self.pdf_var.get(),
            "no_excel": not self.excel_var.get(),
        }
        self.last_out_dir = None

        self._clear_log()
        self._append_log(
            f"Classeur : {workbook}\nAnnée : {params['year']}\n"
            f"PDF : {'oui' if not params['no_pdf'] else 'non'}   "
            f"Excel : {'oui' if not params['no_excel'] else 'non'}\n\n"
        )
        self.run_btn["state"] = "disabled"
        self.open_btn["state"] = "disabled"

        threading.Thread(target=self._worker, args=(params,), daemon=True).start()

    def _worker(self, params):
        # Import here (not at module top) so a missing dependency surfaces in the
        # log instead of crashing the whole app at startup.
        try:
            import generate_factures
        except Exception as exc:  # noqa: BLE001
            self.log_queue.put(("line", f"\nERREUR: impossible de charger le moteur : {exc}\n"))
            self.log_queue.put(("done", -1))
            return

        def log(msg):
            self.log_queue.put(("line", str(msg) + "\n"))

        try:
            out_dir = generate_factures.generate(log=log, **params)
            if out_dir:
                self.last_out_dir = str(out_dir)
            self.log_queue.put(("done", 0))
        except Exception as exc:  # noqa: BLE001
            self.log_queue.put(("line", f"\nERREUR: {exc}\n"))
            self.log_queue.put(("done", -1))

    # -- log plumbing ---------------------------------------------------------
    def _drain_log(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == "line":
                    self._append_log(payload)
                elif kind == "done":
                    self._on_done(payload)
        except queue.Empty:
            pass
        self.root.after(100, self._drain_log)

    def _on_done(self, code):
        self.run_btn["state"] = "normal"
        if code == 0:
            self._append_log("\n✅ Terminé.\n")
            if self.last_out_dir:
                self.open_btn["state"] = "normal"
        else:
            self._append_log(f"\n❌ Échec (code {code}).\n")
            messagebox.showerror("Échec", f"La génération a échoué (code {code}). Voir le journal.")

    def _append_log(self, text):
        self.log["state"] = "normal"
        self.log.insert("end", text)
        self.log.see("end")
        self.log["state"] = "disabled"

    def _clear_log(self):
        self.log["state"] = "normal"
        self.log.delete("1.0", "end")
        self.log["state"] = "disabled"


def main():
    root = tk.Tk()
    FacturesGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
