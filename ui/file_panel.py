import pathlib
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar
from tkinter import ttk
import customtkinter as ctk
import polars as pl
import fastexcel

import operations

# We will move load_file to core.ui_utils
from core.ui_utils import load_file

class FilePanel(ctk.CTkFrame):
    """Browse button + key-column dropdown + 5-row preview table."""

    def __init__(self, master, label: str, on_load=None, **kwargs):
        super().__init__(master, **kwargs)
        self.label    = label
        self.on_load  = on_load
        self.df: pl.DataFrame | None = None
        self.path: str = ""
        self._build()

    def _build(self):
        # ── Header ──────────────────────────────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=8, pady=(8, 2))
        ctk.CTkLabel(hdr, text=self.label, text_color="#cdd6f4",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkButton(hdr, text="Parcourir …", width=110,
                      fg_color="#89b4fa", text_color="#181825", hover_color="#74a1e9",
                      command=self._browse).pack(side="right")

        self._fname_var = StringVar(value="Aucun fichier sélectionné")
        self._fname_lbl = ctk.CTkLabel(self, textvariable=self._fname_var,
                     text_color="#a6adc8",
                     font=ctk.CTkFont(size=11))
        self._fname_lbl.pack(anchor="w", padx=10, pady=(0, 2))

        # ── Sheet selector (shown only for multi-sheet Excel) ────────────────
        self._sheet_row = ctk.CTkFrame(self, fg_color="transparent")
        self._sheet_row.pack(fill="x", padx=8, pady=(0, 2))
        ctk.CTkLabel(self._sheet_row, text="Onglet :").pack(side="left", padx=(0, 6))
        self.sheet_var = StringVar(value="—")
        self.sheet_menu = ctk.CTkOptionMenu(
            self._sheet_row, variable=self.sheet_var,
            values=["—"], width=200, command=self._on_sheet_change)
        self.sheet_menu.pack(side="left", fill="x", expand=True)
        self._sheet_row.pack_forget()          # hidden until multi-sheet file loaded
        self._xls = None

        # ── Key column dropdown ──────────────────────────────────────────────
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", padx=8, pady=(0, 2))
        ctk.CTkLabel(row, text="Colonne clé :").pack(side="left", padx=(0, 6))
        self.col_var  = StringVar(value="— charger un fichier —")
        self.col_menu = ctk.CTkOptionMenu(row, variable=self.col_var,
                                          values=["— charger un fichier —"],
                                          width=200)
        self.col_menu.pack(side="left", fill="x", expand=True)

        self._info_var = StringVar(value="")
        ctk.CTkLabel(self, textvariable=self._info_var,
                     text_color="#a6adc8",
                     font=ctk.CTkFont(size=11)).pack(anchor="w", padx=10)

        # ── Preview table ────────────────────────────────────────────────────
        pf = ctk.CTkFrame(self, fg_color="transparent")
        pf.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        pf.columnconfigure(0, weight=1)
        pf.rowconfigure(0, weight=1)

        self._tree = ttk.Treeview(pf, height=4, show="headings")
        vsb = ttk.Scrollbar(pf, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(pf, orient="horizontal",  command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    def get_column(self) -> str:
        return self.col_var.get()

    def _browse(self):
        path = filedialog.askopenfilename(
            title=f"Sélectionner {self.label}",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tous les fichiers", "*.*")],
        )
        if not path:
            return
        self._fname_var.set("Chargement en cours …")
        self._info_var.set("")

        def _load_in_thread():
            p = pathlib.Path(path)
            xls_sheets = []
            try:
                if p.suffix.lower() in (".xlsx", ".xls"):
                    xls_obj = fastexcel.read_excel(path)
                    xls_sheets = xls_obj.sheet_names
                    # Load first sheet
                    df = operations.cleanup_floats(xls_obj.load_sheet(xls_sheets[0]).to_polars())
                else:
                    df = load_file(path)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Erreur de chargement", str(e)))
                self.after(0, lambda: self._fname_var.set("Aucun fichier sélectionné"))
                return

            def _apply():
                self._xls = None
                self._sheet_row.pack_forget()
                if len(xls_sheets) > 1:
                    self._xls = path
                    self.sheet_menu.configure(values=xls_sheets)
                    self.sheet_var.set(xls_sheets[0])
                    self._sheet_row.pack(fill="x", padx=8, pady=(0, 2),
                                          after=self._fname_lbl)
                self._accept(path, df)
            self.after(0, _apply)

        threading.Thread(target=_load_in_thread, daemon=True).start()

    def _accept(self, path: str, df: pl.DataFrame):
        self.path = path
        self.df   = df
        self._fname_var.set(pathlib.Path(path).name)
        self._info_var.set(f"{len(df):,} lignes · {len(df.columns)} colonnes")
        cols = list(df.columns)
        self.col_menu.configure(values=cols)
        self.col_var.set(cols[0])
        self._populate_preview(df)
        if self.on_load:
            self.on_load()

    def _on_sheet_change(self, sheet_name: str):
        if self._xls is None:
            return
        try:
            df = operations.cleanup_floats(fastexcel.read_excel(self._xls).load_sheet(sheet_name).to_polars())
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
            return
        self.df = df
        self._info_var.set(f"{len(df):,} lignes · {len(df.columns)} colonnes")
        cols = list(df.columns)
        self.col_menu.configure(values=cols)
        self.col_var.set(cols[0])
        self._populate_preview(df)
        if self.on_load:
            self.on_load()

    def _populate_preview(self, df: pl.DataFrame):
        self._tree.delete(*self._tree.get_children())
        self._tree["columns"] = list(df.columns)
        for c in df.columns:
            self._tree.heading(c, text=c)
            self._tree.column(c, width=100, minwidth=50, stretch=False)
        for row in df.head(5).rows():
            self._tree.insert("", "end", values=row)
