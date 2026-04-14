"""
Excel / CSV Filter Tool  — Generic GUI
Replaces the hardcoded filter.py with a fully interactive desktop app.
"""

import datetime
import json
import pathlib
import threading
from tkinter import filedialog, messagebox
from tkinter import ttk

import customtkinter as ctk
import pandas as pd
from tkinter import BooleanVar, StringVar

from core import DataProcessor, load_tabular_file, scan_tabular_source, split_by_columns, split_workbook_sheets, split_by_line_count
from core.file_service import EXCEL_SUFFIXES, SheetMetadata

# ── Defaults ──────────────────────────────────────────────────────────────────

WORKSPACE     = pathlib.Path(__file__).parent
OPERATIONS    = ["Soustraire  (A − B)", "Fusionner  (A ∪ B)", "Intersecter  (A ∩ B)", "Enrichir  (A ← B)", "Combiner  (Multi-fichiers)", "Extraire  (Filtrer A)", "Scinder  (Onglets)", "Scinder  (Colonnes)", "Scinder  (Lignes)"]
CONDITIONS    = [
    "égal à", "différent de", "contient",
    "commence par", "se termine par",
    "supérieur à", "inférieur à",
    "est vide", "n'est pas vide",
]
OP_DESC = {
    "Soustraire  (A − B)":  "Renvoie les lignes de A dont la clé n'est PAS présente dans B.",
    "Fusionner  (A ∪ B)":   "Empile A et B verticalement. Option : supprimer les doublons.",
    "Intersecter  (A ∩ B)": "Renvoie les lignes de A dont la clé EST présente dans B.",
    "Enrichir  (A ← B)":   "Compare A et B sur une clé. Garde toutes les lignes de A et y ajoute les colonnes choisies de B (LEFT JOIN).",
    "Combiner  (Multi-fichiers)": "Empile plusieurs fichiers Excel/CSV, choisissez les colonnes à conserver et dédoublonnez si souhaité.",
    "Extraire  (Filtrer A)": "Filtre les lignes de A selon une condition sur une colonne choisie.",
    "Scinder  (Onglets)":    "Sépare chaque onglet d'un ou plusieurs fichiers Excel en fichiers individuels.",
    "Scinder  (Colonnes)":   "Charge un seul fichier, choisissez les colonnes à conserver et une ou plusieurs colonnes de découpe. Génère un fichier Excel par combinaison de valeurs unique.",
    "Scinder  (Lignes)":     "Charge un seul fichier et le découpe en fichiers plus petits contenant un nombre de lignes spécifié.",
}
                    
# ── Helpers ───────────────────────────────────────────────────────────────────

def load_file(path: str) -> pd.DataFrame:
    return load_tabular_file(path)


def fmt(n) -> str:
    """Format a row-count or return '—' for None."""
    return f"{n:,}" if isinstance(n, int) else "—"


# ── FilePanel ─────────────────────────────────────────────────────────────────

class FilePanel(ctk.CTkFrame):
    """Browse button + key-column dropdown + 5-row preview table."""

    def __init__(self, master, label: str, on_load=None, **kwargs):
        super().__init__(master, **kwargs)
        self.label    = label
        self.on_load  = on_load
        self.df: pd.DataFrame | None = None
        self.path: str = ""
        self._build()

    def _build(self):
        # ── Header ──────────────────────────────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=8, pady=(8, 2))
        ctk.CTkLabel(hdr, text=self.label,
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkButton(hdr, text="Parcourir …", width=110,
                      command=self._browse).pack(side="right")

        self._fname_var = StringVar(value="Aucun fichier sélectionné")
        self._fname_lbl = ctk.CTkLabel(self, textvariable=self._fname_var,
                     text_color="gray",
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
        self._xls: pd.ExcelFile | None = None

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
                     text_color="gray",
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

    # ── Public ───────────────────────────────────────────────────────────────

    def get_column(self) -> str:
        return self.col_var.get()

    # ── Internal ─────────────────────────────────────────────────────────────

    def _browse(self):
        path = filedialog.askopenfilename(
            title=f"Sélectionner {self.label}",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Tous les fichiers", "*.*")],
        )
        if not path:
            return
        self._fname_var.set("Chargement en cours …")
        self._info_var.set("")

        def _load_in_thread():
            p = pathlib.Path(path)
            old_xls = self._xls
            xls_obj = None
            try:
                if p.suffix.lower() in (".xlsx", ".xls"):
                    xls_obj = pd.ExcelFile(path)
                    sheets = xls_obj.sheet_names
                    df = xls_obj.parse(sheets[0])
                else:
                    df = load_file(path)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Erreur de chargement", str(e)))
                self.after(0, lambda: self._fname_var.set("Aucun fichier sélectionné"))
                return

            def _apply():
                # Close previous ExcelFile if any
                if old_xls is not None:
                    try:
                        old_xls.close()
                    except Exception:
                        pass
                self._xls = None
                self._sheet_row.pack_forget()
                if xls_obj is not None and len(xls_obj.sheet_names) > 1:
                    self._xls = xls_obj
                    self.sheet_menu.configure(values=xls_obj.sheet_names)
                    self.sheet_var.set(xls_obj.sheet_names[0])
                    self._sheet_row.pack(fill="x", padx=8, pady=(0, 2),
                                          after=self._fname_lbl)
                elif xls_obj is not None:
                    try:
                        xls_obj.close()
                    except Exception:
                        pass
                self._accept(path, df)
            self.after(0, _apply)

        threading.Thread(target=_load_in_thread, daemon=True).start()

    def _accept(self, path: str, df: pd.DataFrame):
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
        """Re-parse the selected sheet and update preview / columns."""
        if self._xls is None:
            return
        try:
            df = self._xls.parse(sheet_name)
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

    def _populate_preview(self, df: pd.DataFrame):
        self._tree.delete(*self._tree.get_children())
        self._tree["columns"] = list(df.columns)
        for c in df.columns:
            self._tree.heading(c, text=c)
            self._tree.column(c, width=100, minwidth=50, stretch=False)
        for row in df.head(5).values.tolist():
            self._tree.insert("", "end", values=row)


# ── Main Application ──────────────────────────────────────────────────────────

class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Outil de Filtrage Excel / CSV")
        self.geometry("1150x860")
        self.minsize(900, 700)
        self._processor = DataProcessor()
        self._result_df: pd.DataFrame | None = None
        self._notfound_df: pd.DataFrame | None = None
        self._en_found_df: pd.DataFrame | None = None
        self._col_vars:  dict[str, BooleanVar] = {}
        self._build_ui()

    # =========================================================================
    # UI construction
    # =========================================================================

    def _build_ui(self):
        self._build_toolbar()
        self._build_navbar()
        scroll = ctk.CTkScrollableFrame(self)
        scroll.pack(fill="both", expand=True, padx=10, pady=6)
        self._scroll = scroll
        self._build_files_section()
        self._build_operation_section()
        # Run button
        ctk.CTkButton(
            scroll, text="▶  Lancer l'opération", height=44,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._run,
        ).pack(fill="x", padx=4, pady=8)
        self._build_result_section()

    # ── Toolbar ───────────────────────────────────────────────────────────────

    def _build_navbar(self):
        """Fixed operation-selector bar pinned below the toolbar."""
        nav = ctk.CTkFrame(self, corner_radius=0,
                           fg_color=("gray88", "gray18"))
        nav.pack(fill="x", side="top")

        self._op_desc_var = StringVar(value=OP_DESC[OPERATIONS[0]])
        ctk.CTkLabel(
            nav, textvariable=self._op_desc_var,
            text_color="#5b9bd5",
            font=ctk.CTkFont(size=11, slant="italic"),
        ).pack(anchor="w", padx=14, pady=(6, 0))

        self._op_var = StringVar(value=OPERATIONS[0])
        ctk.CTkSegmentedButton(
            nav, values=OPERATIONS,
            variable=self._op_var,
            command=self._op_changed,
        ).pack(fill="x", padx=10, pady=(4, 8))

    def _build_toolbar(self):
        tb = ctk.CTkFrame(self, height=50, corner_radius=0)
        tb.pack(fill="x", side="top")
        tb.pack_propagate(False)

        ctk.CTkLabel(
            tb, text="⚙  Outil de Filtrage Excel / CSV",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(side="left", padx=16)

        # right side (pack right-to-left)
        self._theme_sw = ctk.CTkSwitch(tb, text="Mode clair",
                                        command=self._toggle_theme)
        self._theme_sw.pack(side="right", padx=12)
        ctk.CTkButton(tb, text="Tout effacer", width=105,
                      fg_color="#c0392b", hover_color="#992d22",
                      command=self._clear_all).pack(side="right", padx=4)
        ctk.CTkButton(tb, text="Charger profil", width=115,
                      command=self._load_profile).pack(side="right", padx=4)
        ctk.CTkButton(tb, text="Sauvegarder profil", width=135,
                      command=self._save_profile).pack(side="right", padx=4)

    # ── Section 1: Files ──────────────────────────────────────────────────────

    def _build_files_section(self):
        sec = ctk.CTkFrame(self._scroll)
        sec.pack(fill="x", padx=4, pady=6)
        self._files_sec = sec          # keep reference for hide/show
        ctk.CTkLabel(sec, text="1 · Charger les fichiers",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     ).pack(anchor="w", padx=10, pady=(8, 4))

        self._pa = FilePanel(sec, "Fichier A  (principal)",   on_load=self._file_loaded)
        self._pb = FilePanel(sec, "Fichier B  (secondaire)", on_load=self._file_loaded)
        self._pa.pack(fill="x", padx=6, pady=(0, 4))
        self._pb.pack(fill="x", padx=6, pady=(0, 8))

    # ── Section 2: Operation ──────────────────────────────────────────────────

    def _build_operation_section(self):
        self._op_sec = ctk.CTkFrame(self._scroll)
        self._op_sec.pack(fill="x", padx=4, pady=6)

        host = ctk.CTkFrame(self._op_sec, fg_color="transparent")
        host.pack(fill="x", padx=10, pady=(8, 10))
        self._op_host = host

        self._sub: dict[str, ctk.CTkFrame] = {
            "mi":       self._build_mi_panel(host),
            "extend":   self._build_extend_panel(host),
            "enrich":   self._build_enrich_panel(host),
            "combine":  self._build_combine_panel(host),
            "extract":  self._build_extract_panel(host),
            "split":    self._build_split_panel(host),
            "splitcol": self._build_splitcol_panel(host),
            "splitlines": self._build_splitlines_panel(host),
        }
        self._show_sub("mi")

    def _build_mi_panel(self, parent) -> ctk.CTkFrame:
        f = ctk.CTkFrame(parent, fg_color="transparent")

        # ── Mode selector ───────────────────────────────────────────────
        mode_row = ctk.CTkFrame(f, fg_color="transparent")
        mode_row.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(4, 6))
        ctk.CTkLabel(mode_row, text="Mode :",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 8))
        self._mi_mode_var = StringVar(value="Standard (1 A × 1 B)")
        ctk.CTkOptionMenu(
            mode_row, variable=self._mi_mode_var,
            values=["Standard (1 A × 1 B)",
                    "1 A × Plusieurs B",
                    "Plusieurs A × 1 B"],
            width=240,
            command=self._mi_mode_changed,
        ).pack(side="left")

        # ── Key column dropdowns (standard mode) ────────────────────────
        ctk.CTkLabel(f, text="Colonne clé dans Fichier A :").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self._mi_a_var  = StringVar(value="— charger Fichier A —")
        self._mi_a_menu = ctk.CTkOptionMenu(f, variable=self._mi_a_var,
                                             values=["— charger Fichier A —"], width=220)
        self._mi_a_menu.grid(row=1, column=1, sticky="w", pady=4)

        ctk.CTkLabel(f, text="Colonne clé dans Fichier B :").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self._mi_b_var  = StringVar(value="— charger Fichier B —")
        self._mi_b_menu = ctk.CTkOptionMenu(f, variable=self._mi_b_var,
                                             values=["— charger Fichier B —"], width=220)
        self._mi_b_menu.grid(row=2, column=1, sticky="w", pady=4)

        ctk.CTkLabel(
            f,
            text="Les clés sont normalisées (chaînes sans espaces) avant comparaison.",
            text_color="gray", font=ctk.CTkFont(size=11),
        ).grid(row=3, column=0, columnspan=2, sticky="w")

        # ── Multi-file picker (hidden by default) ───────────────────────
        self._mi_multi_frame = ctk.CTkFrame(f)
        self._mi_multi_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        mi_btn_row = ctk.CTkFrame(self._mi_multi_frame, fg_color="transparent")
        mi_btn_row.pack(fill="x", padx=6, pady=(6, 4))
        self._mi_multi_label = ctk.CTkLabel(
            mi_btn_row, text="Fichiers B supplémentaires :",
            font=ctk.CTkFont(size=12, weight="bold"))
        self._mi_multi_label.pack(side="left")
        ctk.CTkButton(mi_btn_row, text="Ajouter …", width=100,
                      command=self._mi_multi_browse).pack(side="left", padx=(12, 4))
        ctk.CTkButton(mi_btn_row, text="Retirer", width=80,
                      fg_color="#c0392b", hover_color="#992d22",
                      command=self._mi_multi_remove).pack(side="left", padx=4)

        # Key column for multi-files
        mi_key_row = ctk.CTkFrame(self._mi_multi_frame, fg_color="transparent")
        mi_key_row.pack(fill="x", padx=6, pady=(0, 4))
        ctk.CTkLabel(mi_key_row, text="Colonne clé dans ces fichiers :").pack(side="left", padx=(0, 6))
        self._mi_multi_key_var = StringVar(value="— ajouter des fichiers —")
        self._mi_multi_key_menu = ctk.CTkOptionMenu(
            mi_key_row, variable=self._mi_multi_key_var,
            values=["— ajouter des fichiers —"], width=220)
        self._mi_multi_key_menu.pack(side="left")

        mi_list_fr = ctk.CTkFrame(self._mi_multi_frame)
        mi_list_fr.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        mi_list_fr.columnconfigure(0, weight=1)
        mi_list_fr.rowconfigure(0, weight=1)
        self._mi_multi_tree = ttk.Treeview(
            mi_list_fr, columns=("cols", "rows"), height=4, show="headings")
        self._mi_multi_tree.heading("cols", text="Colonnes")
        self._mi_multi_tree.heading("rows", text="Lignes")
        self._mi_multi_tree.column("cols", width=80, anchor="center")
        self._mi_multi_tree.column("rows", width=80, anchor="center")
        mi_vsb = ttk.Scrollbar(mi_list_fr, orient="vertical",
                               command=self._mi_multi_tree.yview)
        self._mi_multi_tree.configure(yscrollcommand=mi_vsb.set)
        self._mi_multi_tree.grid(row=0, column=0, sticky="nsew")
        mi_vsb.grid(row=0, column=1, sticky="ns")

        # Internal storage: path → DataFrame
        self._mi_multi_files: dict[str, pd.DataFrame] = {}

        self._mi_multi_frame.grid_remove()     # hidden until multi mode

        # ── "Non trouvés" checkbox (intersect only) ─────────────────────
        self._mi_nf_var = BooleanVar(value=True)
        ctk.CTkCheckBox(
            f, text="Inclure les non trouvés (lignes de B absentes dans A)",
            variable=self._mi_nf_var,
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # ── Column selector (intersect only) ────────────────────────────
        self._mi_col_frame = ctk.CTkFrame(f)
        self._mi_col_frame.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        ctk.CTkLabel(self._mi_col_frame,
                     text="Colonnes à inclure dans le résultat :",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     ).pack(anchor="w", padx=6, pady=(6, 4))

        cols_host = ctk.CTkFrame(self._mi_col_frame, fg_color="transparent")
        cols_host.pack(fill="x", padx=4, pady=(0, 6))
        cols_host.columnconfigure(0, weight=1)
        cols_host.columnconfigure(1, weight=1)

        # -- File A column list --
        a_fr = ctk.CTkFrame(cols_host)
        a_fr.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        a_hdr = ctk.CTkFrame(a_fr, fg_color="transparent")
        a_hdr.pack(fill="x", padx=4, pady=(4, 2))
        ctk.CTkLabel(a_hdr, text="Fichier A",
                     font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")
        ctk.CTkButton(a_hdr, text="Rien", width=50, height=22,
                      command=lambda: self._mi_toggle_all("a", False)).pack(side="right", padx=2)
        ctk.CTkButton(a_hdr, text="Tout", width=50, height=22,
                      command=lambda: self._mi_toggle_all("a", True)).pack(side="right", padx=2)
        self._mi_a_col_scroll = ctk.CTkScrollableFrame(a_fr, height=120)
        self._mi_a_col_scroll.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        # -- File B column list --
        b_fr = ctk.CTkFrame(cols_host)
        b_fr.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        b_hdr = ctk.CTkFrame(b_fr, fg_color="transparent")
        b_hdr.pack(fill="x", padx=4, pady=(4, 2))
        ctk.CTkLabel(b_hdr, text="Fichier B",
                     font=ctk.CTkFont(size=11, weight="bold")).pack(side="left")
        ctk.CTkButton(b_hdr, text="Rien", width=50, height=22,
                      command=lambda: self._mi_toggle_all("b", False)).pack(side="right", padx=2)
        ctk.CTkButton(b_hdr, text="Tout", width=50, height=22,
                      command=lambda: self._mi_toggle_all("b", True)).pack(side="right", padx=2)
        self._mi_b_col_scroll = ctk.CTkScrollableFrame(b_fr, height=120)
        self._mi_b_col_scroll.pack(fill="both", expand=True, padx=4, pady=(0, 4))

        self._mi_a_col_vars: dict[str, BooleanVar] = {}
        self._mi_b_col_vars: dict[str, BooleanVar] = {}

        # Principal source + suffix for duplicate names
        sfx_row = ctk.CTkFrame(self._mi_col_frame, fg_color="transparent")
        sfx_row.pack(fill="x", padx=6, pady=(2, 6))
        ctk.CTkLabel(sfx_row, text="Colonnes principales :",
                 font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
        self._mi_principal_var = StringVar(value="A")
        ctk.CTkOptionMenu(sfx_row, variable=self._mi_principal_var,
                  values=["A", "B"], width=70).pack(side="left")
        ctk.CTkLabel(sfx_row, text="   Suffixe côté secondaire :",
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
        self._mi_suffix_var = StringVar(value=" (B)")
        ctk.CTkEntry(sfx_row, textvariable=self._mi_suffix_var, width=120,
                     placeholder_text=" (B)").pack(side="left")
        ctk.CTkLabel(sfx_row, text="  Appliqué au fichier non principal en cas de conflit",
                     text_color="gray", font=ctk.CTkFont(size=10)).pack(side="left", padx=8)

        # Hidden by default \u2013 shown only for Intersecter
        self._mi_col_frame.grid_remove()

        return f

    def _build_extend_panel(self, parent) -> ctk.CTkFrame:
        f = ctk.CTkFrame(parent, fg_color="transparent")
        self._ext_dedup = BooleanVar(value=True)
        ctk.CTkCheckBox(f, text="Supprimer les doublons après fusion",
                         variable=self._ext_dedup).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=4)

        ctk.CTkLabel(f, text="Colonne clé pour doublons :").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self._ext_key_var  = StringVar(value="— aucune (toutes les colonnes) —")
        self._ext_key_menu = ctk.CTkOptionMenu(
            f, variable=self._ext_key_var,
            values=["— aucune (toutes les colonnes) —"], width=240)
        self._ext_key_menu.grid(row=1, column=1, sticky="w", pady=4)
        return f

    def _build_extract_panel(self, parent) -> ctk.CTkFrame:
        f = ctk.CTkFrame(parent, fg_color="transparent")

        ctk.CTkLabel(f, text="Colonne à filtrer :").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self._ex_col_var  = StringVar(value="— charger Fichier A —")
        self._ex_col_menu = ctk.CTkOptionMenu(
            f, variable=self._ex_col_var,
            values=["— charger Fichier A —"], width=220)
        self._ex_col_menu.grid(row=0, column=1, sticky="w", pady=4)

        ctk.CTkLabel(f, text="Condition :").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self._ex_cond_var = StringVar(value=CONDITIONS[0])
        self._ex_cond_menu = ctk.CTkOptionMenu(
            f, variable=self._ex_cond_var, values=CONDITIONS,
            command=self._cond_changed, width=180)
        self._ex_cond_menu.grid(row=1, column=1, sticky="w", pady=4)

        ctk.CTkLabel(f, text="Valeur :").grid(
            row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self._ex_val_var  = StringVar()
        self._ex_val_entry = ctk.CTkEntry(
            f, textvariable=self._ex_val_var,
            width=220, placeholder_text="saisir une valeur …")
        self._ex_val_entry.grid(row=2, column=1, sticky="w", pady=4)

        self._ex_case_var = BooleanVar(value=False)
        ctk.CTkCheckBox(f, text="Sensible à la casse",
                         variable=self._ex_case_var).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=4)
        return f

    def _build_enrich_panel(self, parent) -> ctk.CTkFrame:
        """Panel for Enrichir (A ← B): left-join A with selected B columns."""
        f = ctk.CTkFrame(parent, fg_color="transparent")

        # ── Key column dropdowns ────────────────────────────────────────
        ctk.CTkLabel(f, text="Colonne clé dans Fichier A :").grid(
            row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self._en_a_var  = StringVar(value="— charger Fichier A —")
        self._en_a_menu = ctk.CTkOptionMenu(f, variable=self._en_a_var,
                                             values=["— charger Fichier A —"], width=220)
        self._en_a_menu.grid(row=0, column=1, sticky="w", pady=4)

        ctk.CTkLabel(f, text="Colonne clé dans Fichier B :").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self._en_b_var  = StringVar(value="— charger Fichier B —")
        self._en_b_menu = ctk.CTkOptionMenu(f, variable=self._en_b_var,
                                             values=["— charger Fichier B —"], width=220)
        self._en_b_menu.grid(row=1, column=1, sticky="w", pady=4)

        ctk.CTkLabel(
            f,
            text="Les clés sont normalisées (chaînes sans espaces) avant comparaison.",
            text_color="gray", font=ctk.CTkFont(size=11),
        ).grid(row=2, column=0, columnspan=2, sticky="w")

        # ── Join mode ───────────────────────────────────────────────────
        mode_row = ctk.CTkFrame(f, fg_color="transparent")
        mode_row.grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 2))
        ctk.CTkLabel(mode_row, text="Mode :").pack(side="left", padx=(0, 6))
        self._en_mode_var = StringVar(value="Toutes les lignes de A (LEFT JOIN)")
        ctk.CTkOptionMenu(mode_row, variable=self._en_mode_var,
                          values=[
                              "Toutes les lignes de A (LEFT JOIN)",
                              "Seulement les correspondances (INNER JOIN)",
                          ], width=340).pack(side="left")

        # ── B columns to bring in ──────────────────────────────────────
        sel_frame = ctk.CTkFrame(f)
        sel_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        sel_hdr = ctk.CTkFrame(sel_frame, fg_color="transparent")
        sel_hdr.pack(fill="x", padx=6, pady=(6, 4))
        ctk.CTkLabel(sel_hdr, text="Colonnes de B à ajouter au résultat :",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")
        ctk.CTkButton(sel_hdr, text="Rien", width=50, height=22,
                      command=lambda: self._en_toggle_all(False)).pack(side="right", padx=2)
        ctk.CTkButton(sel_hdr, text="Tout", width=50, height=22,
                      command=lambda: self._en_toggle_all(True)).pack(side="right", padx=2)

        self._en_col_scroll = ctk.CTkScrollableFrame(sel_frame, height=130)
        self._en_col_scroll.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        self._en_col_vars: dict[str, BooleanVar] = {}

        # ── Suffix for duplicate names ─────────────────────────────────
        sfx_row = ctk.CTkFrame(sel_frame, fg_color="transparent")
        sfx_row.pack(fill="x", padx=6, pady=(2, 6))
        ctk.CTkLabel(sfx_row, text="Suffixe si conflit de nom :",
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
        self._en_suffix_var = StringVar(value=" (B)")
        ctk.CTkEntry(sfx_row, textvariable=self._en_suffix_var, width=120,
                     placeholder_text=" (B)").pack(side="left")
        ctk.CTkLabel(sfx_row, text="  Ajouté aux colonnes de B en cas de doublon",
                     text_color="gray", font=ctk.CTkFont(size=10)).pack(side="left", padx=8)

        # ── Empty-cell handling ───────────────────────────────────────────
        empty_frame = ctk.CTkFrame(f)
        empty_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        self._en_skip_empty = BooleanVar(value=True)
        ctk.CTkCheckBox(empty_frame,
                        text="Ignorer les lignes dont la clé est vide / NaN",
                        variable=self._en_skip_empty
                        ).pack(anchor="w", padx=6, pady=(6, 4))

        fill_row = ctk.CTkFrame(empty_frame, fg_color="transparent")
        fill_row.pack(fill="x", padx=6, pady=(2, 6))
        ctk.CTkLabel(fill_row, text="Cellules sans correspondance :",
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
        self._en_fill_var = StringVar(value="Laisser vide")
        ctk.CTkOptionMenu(fill_row, variable=self._en_fill_var,
                          values=["Laisser vide", "Chaîne vide", "Valeur personnalisée"],
                          width=190,
                          command=self._en_fill_changed).pack(side="left")
        self._en_fill_entry = ctk.CTkEntry(fill_row, width=140,
                                            placeholder_text="valeur …")
        self._en_fill_entry.pack(side="left", padx=(8, 0))
        self._en_fill_entry.configure(state="disabled")

        return f

    def _build_split_panel(self, parent) -> ctk.CTkFrame:
        f = ctk.CTkFrame(parent, fg_color="transparent")

        # ── File selection ───────────────────────────────────────────────
        file_row = ctk.CTkFrame(f, fg_color="transparent")
        file_row.pack(fill="x", pady=(4, 4))
        ctk.CTkButton(file_row, text="Ajouter des fichiers …", width=170,
                      command=self._split_browse).pack(side="left")
        ctk.CTkButton(file_row, text="Retirer sélection", width=130,
                      fg_color="#c0392b", hover_color="#992d22",
                      command=self._split_remove).pack(side="left", padx=8)

        # File list with sheet details
        list_fr = ctk.CTkFrame(f)
        list_fr.pack(fill="both", expand=True, pady=(0, 4))
        list_fr.columnconfigure(0, weight=1)
        list_fr.rowconfigure(0, weight=1)
        self._sp_tree = ttk.Treeview(
            list_fr, columns=("sheets", "rows"), height=6, show="tree headings",
        )
        self._sp_tree.heading("#0", text="Fichier")
        self._sp_tree.heading("sheets", text="Onglets")
        self._sp_tree.heading("rows", text="Lignes")
        self._sp_tree.column("#0", width=300, stretch=True)
        self._sp_tree.column("sheets", width=80, anchor="center")
        self._sp_tree.column("rows", width=80, anchor="center")
        sp_vsb = ttk.Scrollbar(list_fr, orient="vertical", command=self._sp_tree.yview)
        self._sp_tree.configure(yscrollcommand=sp_vsb.set)
        self._sp_tree.grid(row=0, column=0, sticky="nsew")
        sp_vsb.grid(row=0, column=1, sticky="ns")

        # Internal storage: path → {sheet_name: metadata, …}
        self._sp_files: dict[str, dict[str, SheetMetadata]] = {}

        # ── Naming ───────────────────────────────────────────────────────
        nm = ctk.CTkFrame(f, fg_color="transparent")
        nm.pack(fill="x", pady=(4, 4))
        ctk.CTkLabel(nm, text="Nommage :").pack(side="left", padx=(0, 6))
        self._sp_name_var = StringVar(value="{fichier}_{onglet}")
        ctk.CTkEntry(nm, textvariable=self._sp_name_var, width=300,
                     placeholder_text="{fichier}_{onglet}").pack(side="left")
        ctk.CTkLabel(nm, text="  Placeholders : {fichier}, {onglet}",
                     text_color="gray", font=ctk.CTkFont(size=11)).pack(side="left", padx=8)

        # ── Format + output dir ──────────────────────────────────────────
        out_row = ctk.CTkFrame(f, fg_color="transparent")
        out_row.pack(fill="x", pady=(4, 4))
        ctk.CTkLabel(out_row, text="Format :").pack(side="left", padx=(0, 6))
        self._sp_fmt_var = StringVar(value="xlsx")
        ctk.CTkOptionMenu(out_row, variable=self._sp_fmt_var,
                          values=["xlsx", "csv"], width=90).pack(side="left")
        ctk.CTkLabel(out_row, text="   Dossier de sortie :").pack(side="left", padx=(16, 6))
        self._sp_dir_var = StringVar()
        ctk.CTkEntry(out_row, textvariable=self._sp_dir_var, width=260,
                     placeholder_text="(même dossier que le fichier source)").pack(side="left")
        ctk.CTkButton(out_row, text="…", width=36,
                      command=self._split_pick_dir).pack(side="left", padx=4)

        return f

    # ── Splitcol panel ────────────────────────────────────────────────────────

    def _build_splitcol_panel(self, parent) -> ctk.CTkFrame:
        """Panel for Scinder (Colonnes): single-file column-based split."""
        f = ctk.CTkFrame(parent, fg_color="transparent")

        # ── File picker row ──────────────────────────────────────────────────
        file_row = ctk.CTkFrame(f, fg_color="transparent")
        file_row.pack(fill="x", pady=(4, 2))
        ctk.CTkButton(file_row, text="Parcourir …", width=130,
                      command=self._sc_browse).pack(side="left")
        self._sc_fname_var = StringVar(value="Aucun fichier sélectionné")
        ctk.CTkLabel(file_row, textvariable=self._sc_fname_var,
                     text_color="gray", font=ctk.CTkFont(size=11)
                     ).pack(side="left", padx=10)

        # ── Sheet selector (hidden if single-sheet) ──────────────────────────
        self._sc_sheet_row = ctk.CTkFrame(f, fg_color="transparent")
        self._sc_sheet_row.pack(fill="x", pady=(0, 2))
        ctk.CTkLabel(self._sc_sheet_row, text="Onglet :").pack(side="left", padx=(0, 6))
        self._sc_sheet_var = StringVar(value="—")
        self._sc_sheet_menu = ctk.CTkOptionMenu(
            self._sc_sheet_row, variable=self._sc_sheet_var,
            values=["—"], width=220, command=self._sc_on_sheet_change)
        self._sc_sheet_menu.pack(side="left")
        self._sc_sheet_row.pack_forget()
        self._sc_xls = None
        self._sc_df: pd.DataFrame | None = None

        # ── All-sheets option ────────────────────────────────────────────────
        self._sc_all_sheets_var = BooleanVar(value=False)
        self._sc_all_sheets_chk = ctk.CTkCheckBox(
            f,
            text="Traiter tous les onglets  (appliquer la même configuration à chaque onglet)",
            variable=self._sc_all_sheets_var,
            command=self._sc_toggle_all_sheets,
        )
        self._sc_all_sheets_chk.pack(anchor="w", padx=8, pady=(2, 4))
        self._sc_all_sheets_chk.pack_forget()   # shown only for multi-sheet files

        # ── Two-column layout: Keep cols (left) | Split-by cols (right) ──────
        two_col = ctk.CTkFrame(f, fg_color="transparent")
        two_col.pack(fill="both", expand=True, pady=(6, 4))
        self._sc_two_col = two_col
        two_col.columnconfigure(0, weight=1)
        two_col.columnconfigure(1, weight=1)

        # -- LEFT: columns to KEEP --
        keep_fr = ctk.CTkFrame(two_col)
        keep_fr.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        keep_hdr = ctk.CTkFrame(keep_fr, fg_color="transparent")
        keep_hdr.pack(fill="x", padx=4, pady=(4, 2))
        ctk.CTkLabel(keep_hdr, text="Colonnes à conserver",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")
        ctk.CTkButton(keep_hdr, text="Tout", width=46, height=22,
                      command=lambda: self._sc_toggle_keep(True)).pack(side="right", padx=2)
        ctk.CTkButton(keep_hdr, text="Rien", width=46, height=22,
                      command=lambda: self._sc_toggle_keep(False)).pack(side="right", padx=2)
        self._sc_keep_scroll = ctk.CTkScrollableFrame(keep_fr, height=180)
        self._sc_keep_scroll.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        self._sc_keep_vars: dict[str, BooleanVar] = {}

        # -- RIGHT: split-by ordered columns --
        by_fr = ctk.CTkFrame(two_col)
        by_fr.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        by_hdr = ctk.CTkFrame(by_fr, fg_color="transparent")
        by_hdr.pack(fill="x", padx=4, pady=(4, 2))
        ctk.CTkLabel(by_hdr, text="Colonnes de découpe  (ordre = hiérarchie)",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")

        # Available columns dropdown + Add button
        avail_row = ctk.CTkFrame(by_fr, fg_color="transparent")
        avail_row.pack(fill="x", padx=4, pady=(2, 2))
        self._sc_avail_var = StringVar(value="— charger un fichier —")
        self._sc_avail_menu = ctk.CTkOptionMenu(
            avail_row, variable=self._sc_avail_var,
            values=["— charger un fichier —"], width=200)
        self._sc_avail_menu.pack(side="left")
        ctk.CTkButton(avail_row, text="Ajouter ▼", width=90,
                      command=self._sc_add_split_col).pack(side="left", padx=(6, 0))

        # Ordered list of chosen split-by columns
        list_fr = ctk.CTkFrame(by_fr)
        list_fr.pack(fill="both", expand=True, padx=4, pady=(4, 4))
        list_fr.columnconfigure(0, weight=1)
        list_fr.rowconfigure(0, weight=1)
        self._sc_split_list = ttk.Treeview(list_fr, columns=("col",), height=8,
                                            show="headings", selectmode="browse")
        self._sc_split_list.heading("col", text="Colonne de découpe (dans l'ordre)")
        self._sc_split_list.column("col", width=200, stretch=True)
        sc_vsb = ttk.Scrollbar(list_fr, orient="vertical",
                               command=self._sc_split_list.yview)
        self._sc_split_list.configure(yscrollcommand=sc_vsb.set)
        self._sc_split_list.grid(row=0, column=0, sticky="nsew")
        sc_vsb.grid(row=0, column=1, sticky="ns")
        self._sc_split_cols: list[str] = []   # ordered

        # Buttons to reorder / remove
        ctrl_row = ctk.CTkFrame(by_fr, fg_color="transparent")
        ctrl_row.pack(fill="x", padx=4, pady=(0, 4))
        ctk.CTkButton(ctrl_row, text="↑  Monter", width=90,
                      command=self._sc_move_up).pack(side="left", padx=(0, 4))
        ctk.CTkButton(ctrl_row, text="↓  Descendre", width=110,
                      command=self._sc_move_down).pack(side="left", padx=(0, 4))
        ctk.CTkButton(ctrl_row, text="✕  Retirer", width=90,
                      fg_color="#c0392b", hover_color="#992d22",
                      command=self._sc_remove_split_col).pack(side="left")

        # ── Output options row ───────────────────────────────────────────────
        opt_row = ctk.CTkFrame(f, fg_color="transparent")
        opt_row.pack(fill="x", pady=(2, 2))
        ctk.CTkLabel(opt_row, text="Dossier de sortie :").pack(side="left", padx=(0, 6))
        self._sc_dir_var = StringVar()
        ctk.CTkEntry(opt_row, textvariable=self._sc_dir_var, width=300,
                     placeholder_text="(même dossier que le fichier source)").pack(side="left")
        ctk.CTkButton(opt_row, text="…", width=36,
                      command=self._sc_pick_dir).pack(side="left", padx=4)

        # ── Zebra mode ───────────────────────────────────────────────────────
        zebra_row = ctk.CTkFrame(f, fg_color="transparent")
        zebra_row.pack(fill="x", pady=(2, 4))
        self._sc_zebra_var = BooleanVar(value=False)
        ctk.CTkCheckBox(
            zebra_row,
            text="Mode zébré  (colorer les lignes en alternance dans les fichiers de sortie)",
            variable=self._sc_zebra_var,
        ).pack(side="left")

        # ── Progress bar (hidden until a run starts) ─────────────────────────
        self._sc_prog_frame = ctk.CTkFrame(f, fg_color="transparent")
        self._sc_prog_frame.pack(fill="x", padx=4, pady=(2, 4))
        self._sc_prog_label = ctk.CTkLabel(
            self._sc_prog_frame, text="",
            font=ctk.CTkFont(size=11))
        self._sc_prog_label.pack(anchor="w", padx=4)
        self._sc_prog_bar = ctk.CTkProgressBar(self._sc_prog_frame, height=14)
        self._sc_prog_bar.set(0)
        self._sc_prog_bar.pack(fill="x", padx=4, pady=(2, 4))
        self._sc_prog_frame.pack_forget()

        return f

    # ── Splitcol helpers ──────────────────────────────────────────────────────

    def _sc_browse(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier à scinder",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Tous", "*.*")],
        )
        if not path:
            return
        self._sc_fname_var.set("Chargement …")

        def _load():
            p = pathlib.Path(path)
            old_xls = self._sc_xls
            xls_obj = None
            try:
                if p.suffix.lower() in (".xlsx", ".xls"):
                    xls_obj = pd.ExcelFile(path)
                    sheets = xls_obj.sheet_names
                    df = xls_obj.parse(sheets[0])
                else:
                    df = load_file(path)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
                self.after(0, lambda: self._sc_fname_var.set("Aucun fichier sélectionné"))
                return

            def _apply():
                if old_xls is not None:
                    try:
                        old_xls.close()
                    except Exception:
                        pass
                self._sc_xls = None
                self._sc_sheet_row.pack_forget()
                self._sc_all_sheets_chk.pack_forget()
                if xls_obj is not None and len(xls_obj.sheet_names) > 1:
                    self._sc_xls = xls_obj
                    self._sc_sheet_menu.configure(values=xls_obj.sheet_names)
                    self._sc_sheet_var.set(xls_obj.sheet_names[0])
                    self._sc_sheet_row.pack(fill="x", pady=(0, 2), before=self._sc_two_col)
                    self._sc_all_sheets_var.set(False)
                    self._sc_all_sheets_chk.pack(anchor="w", padx=8, pady=(2, 4),
                                                  before=self._sc_two_col)
                elif xls_obj is not None:
                    try:
                        xls_obj.close()
                    except Exception:
                        pass
                self._sc_accept(path, df)
            self.after(0, _apply)

        threading.Thread(target=_load, daemon=True).start()

    def _sc_accept(self, path: str, df: pd.DataFrame):
        self._sc_df = df
        self._sc_source_path = path
        self._sc_fname_var.set(pathlib.Path(path).name)
        self._sc_split_cols.clear()
        self._sc_split_list.delete(*self._sc_split_list.get_children())
        self._sc_rebuild_cols()

    def _sc_toggle_all_sheets(self):
        """Show/hide the single-sheet selector depending on the all-sheets toggle."""
        if self._sc_all_sheets_var.get():
            self._sc_sheet_row.pack_forget()
        else:
            if self._sc_xls is not None:
                self._sc_sheet_row.pack(fill="x", pady=(0, 2),
                                        before=self._sc_two_col)

    def _sc_on_sheet_change(self, sheet_name: str):
        if self._sc_xls is None:
            return
        try:
            df = self._sc_xls.parse(sheet_name)
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
            return
        self._sc_df = df
        self._sc_split_cols.clear()
        self._sc_split_list.delete(*self._sc_split_list.get_children())
        self._sc_rebuild_cols()

    def _sc_rebuild_cols(self):
        """Rebuild the keep-checkbox list and available-column dropdown."""
        # Clear keep list
        for w in self._sc_keep_scroll.winfo_children():
            w.destroy()
        self._sc_keep_vars.clear()
        if self._sc_df is None:
            self._sc_avail_menu.configure(values=["— charger un fichier —"])
            self._sc_avail_var.set("— charger un fichier —")
            return
        cols = list(self._sc_df.columns)
        for col in cols:
            v = BooleanVar(value=True)
            ctk.CTkCheckBox(self._sc_keep_scroll, text=str(col),
                            variable=v).pack(anchor="w", padx=4, pady=1)
            self._sc_keep_vars[str(col)] = v
        self._sc_avail_menu.configure(values=cols)
        self._sc_avail_var.set(cols[0] if cols else "")

    def _sc_toggle_keep(self, state: bool):
        for v in self._sc_keep_vars.values():
            v.set(state)

    def _sc_add_split_col(self):
        col = self._sc_avail_var.get()
        if not col or col.startswith("—"):
            return
        if col in self._sc_split_cols:
            messagebox.showwarning("Déjà ajouté",
                                   f"La colonne '{col}' est déjà dans la liste de découpe.")
            return
        self._sc_split_cols.append(col)
        self._sc_split_list.insert("", "end", iid=col, values=(col,))

    def _sc_remove_split_col(self):
        sel = self._sc_split_list.selection()
        if not sel:
            return
        iid = sel[0]
        self._sc_split_list.delete(iid)
        if iid in self._sc_split_cols:
            self._sc_split_cols.remove(iid)

    def _sc_move_up(self):
        sel = self._sc_split_list.selection()
        if not sel:
            return
        iid = sel[0]
        idx = self._sc_split_list.index(iid)
        if idx == 0:
            return
        self._sc_split_list.move(iid, "", idx - 1)
        # keep internal list in sync
        self._sc_split_cols.insert(idx - 1, self._sc_split_cols.pop(idx))

    def _sc_move_down(self):
        sel = self._sc_split_list.selection()
        if not sel:
            return
        iid = sel[0]
        idx = self._sc_split_list.index(iid)
        last = len(self._sc_split_list.get_children()) - 1
        if idx >= last:
            return
        self._sc_split_list.move(iid, "", idx + 1)
        self._sc_split_cols.insert(idx + 1, self._sc_split_cols.pop(idx))

    def _sc_pick_dir(self):
        d = filedialog.askdirectory(title="Dossier de sortie")
        if d:
            self._sc_dir_var.set(d)

    # ── Split lines panel ─────────────────────────────────────────────────────

    def _build_splitlines_panel(self, parent) -> ctk.CTkFrame:
        """Panel for Scinder (Lignes): split a single file by exact line count."""
        f = ctk.CTkFrame(parent, fg_color="transparent")

        # ── File picker row ──────────────────────────────────────────────────
        file_row = ctk.CTkFrame(f, fg_color="transparent")
        file_row.pack(fill="x", pady=(4, 6))
        ctk.CTkButton(file_row, text="Parcourir …", width=130,
                      command=self._sl_browse).pack(side="left")
        self._sl_fname_var = StringVar(value="Aucun fichier sélectionné")
        ctk.CTkLabel(file_row, textvariable=self._sl_fname_var,
                     text_color="gray", font=ctk.CTkFont(size=11)
                     ).pack(side="left", padx=10)

        self._sl_df: pd.DataFrame | None = None
        self._sl_xls = None

        # ── Settings row ──────────────────────────────────────────────────────
        set_row = ctk.CTkFrame(f, fg_color="transparent")
        set_row.pack(fill="x", pady=(4, 6))
        
        ctk.CTkLabel(set_row, text="Lignes par fichier :").pack(side="left", padx=(0, 6))
        self._sl_num_var = StringVar(value="200")
        ctk.CTkEntry(set_row, textvariable=self._sl_num_var, width=80).pack(side="left")
        ctk.CTkLabel(set_row, text="  Le nombre de lignes de données",
                     text_color="gray", font=ctk.CTkFont(size=11)).pack(side="left", padx=8)

        # ── Output options row ───────────────────────────────────────────────
        opt_row = ctk.CTkFrame(f, fg_color="transparent")
        opt_row.pack(fill="x", pady=(4, 6))
        ctk.CTkLabel(opt_row, text="Dossier de sortie :").pack(side="left", padx=(0, 6))
        self._sl_dir_var = StringVar()
        ctk.CTkEntry(opt_row, textvariable=self._sl_dir_var, width=300,
                     placeholder_text="(même dossier que le fichier source)").pack(side="left")
        ctk.CTkButton(opt_row, text="…", width=36,
                      command=self._sl_pick_dir).pack(side="left", padx=4)

        # ── Progress bar ─────────────────────────────────────────────────────
        self._sl_prog_frame = ctk.CTkFrame(f, fg_color="transparent")
        self._sl_prog_frame.pack(fill="x", padx=4, pady=(2, 4))
        self._sl_prog_label = ctk.CTkLabel(
            self._sl_prog_frame, text="",
            font=ctk.CTkFont(size=11))
        self._sl_prog_label.pack(anchor="w", padx=4)
        self._sl_prog_bar = ctk.CTkProgressBar(self._sl_prog_frame, height=14)
        self._sl_prog_bar.set(0)
        self._sl_prog_bar.pack(fill="x", padx=4, pady=(2, 4))
        self._sl_prog_frame.pack_forget()

        return f

    def _sl_browse(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier à scinder",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Tous", "*.*")],
        )
        if not path:
            return
        self._sl_fname_var.set("Chargement …")

        def _load():
            p = pathlib.Path(path)
            old_xls = self._sl_xls
            xls_obj = None
            try:
                if p.suffix.lower() in (".xlsx", ".xls"):
                    xls_obj = pd.ExcelFile(path)
                    sheets = xls_obj.sheet_names
                    df = xls_obj.parse(sheets[0])
                else:
                    df = load_file(path)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
                self.after(0, lambda: self._sl_fname_var.set("Aucun fichier sélectionné"))
                return

            def _apply():
                if old_xls is not None:
                    try:
                        old_xls.close()
                    except Exception:
                        pass
                self._sl_xls = xls_obj
                self._sl_df = df
                self._sl_source_path = path
                self._sl_fname_var.set(pathlib.Path(path).name)
            self.after(0, _apply)

        threading.Thread(target=_load, daemon=True).start()

    def _sl_pick_dir(self):
        d = filedialog.askdirectory(title="Dossier de sortie")
        if d:
            self._sl_dir_var.set(d)

    def _op_splitlines(self):
        if self._sl_df is None:
            messagebox.showwarning("Aucun fichier", "Chargez un fichier dans Scinder (Lignes).")
            return

        try:
            lines_per_file = int(self._sl_num_var.get())
            if lines_per_file <= 0:
                raise ValueError()
        except ValueError:
            messagebox.showwarning("Nombre invalide", "Veuillez saisir un nombre entier positif.")
            return

        out_dir_str = self._sl_dir_var.get().strip()
        src_path = getattr(self, "_sl_source_path", "") or ""
        out_dir = out_dir_str if out_dir_str else None

        self._sl_prog_frame.pack(fill="x", padx=4, pady=(2, 4))
        self._sl_prog_bar.set(0)
        self._sl_prog_label.configure(text="Préparation …")
        self.update_idletasks()

        def _progress(done: int, total: int, message: str):
            self.after(0, lambda: (
                self._sl_prog_bar.set(done / total if total else 0),
                self._sl_prog_label.configure(text=message),
            ))

        def _worker():
            total_written = 0
            error_msg = None
            try:
                total_written = split_by_line_count(
                    source_path=src_path,
                    dataframe=self._sl_df,
                    lines_per_file=lines_per_file,
                    output_dir=out_dir,
                    progress_callback=_progress,
                )
            except Exception as e:
                error_msg = str(e)

            def _done():
                if error_msg:
                    self._sl_prog_label.configure(text=f"Erreur : {error_msg}")
                    messagebox.showerror("Erreur de scission (lignes)", error_msg)
                else:
                    self._sl_prog_bar.set(1)
                    summary = f"{total_written} fichier(s) créé(s) de {lines_per_file} lignes (max)"
                    self._sl_prog_label.configure(text=f"✓ {summary}")
                    self._stats_var.set(summary)
                    self._log(f"[Scinder Lignes]  {total_written} fichier(s) → {out_dir or pathlib.Path(src_path).parent}")
                    messagebox.showinfo("Scission terminée", summary)
            self.after(0, _done)

        threading.Thread(target=_worker, daemon=True).start()

    # ── Split helpers ─────────────────────────────────────────────────────────

    def _split_browse(self):
        paths = filedialog.askopenfilenames(
            title="Sélectionner un ou plusieurs fichiers Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tous", "*.*")],
        )
        if not paths:
            return
        for p in paths:
            if p in self._sp_files:
                continue
            try:
                self._sp_files[p] = scan_tabular_source(p)
            except Exception as e:
                messagebox.showerror("Erreur", f"{pathlib.Path(p).name}\n{e}")
        self._split_refresh_tree()

    def _split_remove(self):
        sel = self._sp_tree.selection()
        if not sel:
            return
        for iid in sel:
            # iid is the file path (top-level item)
            parent = self._sp_tree.parent(iid)
            path = iid if not parent else parent
            self._sp_files.pop(path, None)
        self._split_refresh_tree()

    def _split_refresh_tree(self):
        self._sp_tree.delete(*self._sp_tree.get_children())
        for path, sheets in self._sp_files.items():
            total = sum(meta.row_count for meta in sheets.values())
            fid = self._sp_tree.insert(
                "", "end", iid=path,
                text=pathlib.Path(path).name,
                values=(len(sheets), f"{total:,}"),
            )
            for sname, meta in sheets.items():
                self._sp_tree.insert(fid, "end", text=f"  ↳ {sname}",
                                      values=("", f"{meta.row_count:,}"))
            self._sp_tree.item(fid, open=True)

    def _split_pick_dir(self):
        d = filedialog.askdirectory(title="Dossier de sortie")
        if d:
            self._sp_dir_var.set(d)

    # ── Combine helpers ───────────────────────────────────────────────────

    def _combine_browse(self):
        paths = filedialog.askopenfilenames(
            title="Sélectionner des fichiers à combiner",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Tous", "*.*")],
        )
        if not paths:
            return
        for p in paths:
            if p in self._cb_files:
                continue
            try:
                self._cb_files[p] = scan_tabular_source(p)
            except Exception as e:
                messagebox.showerror("Erreur", f"{pathlib.Path(p).name}\n{e}")
        self._combine_refresh_tree()
        self._combine_refresh_cols()

    def _combine_remove(self):
        sel = self._cb_tree.selection()
        if not sel:
            return
        for iid in sel:
            parent = self._cb_tree.parent(iid)
            path = iid if not parent else parent
            self._cb_files.pop(path, None)
        self._combine_refresh_tree()
        self._combine_refresh_cols()

    def _combine_refresh_tree(self):
        self._cb_tree.delete(*self._cb_tree.get_children())
        for path, sheets in self._cb_files.items():
            total = sum(meta.row_count for meta in sheets.values())
            fid = self._cb_tree.insert(
                "", "end", iid=path,
                text=pathlib.Path(path).name,
                values=(len(sheets), f"{total:,}"),
            )
            for sname, meta in sheets.items():
                self._cb_tree.insert(fid, "end",
                                     text=f"  ↳ {sname}",
                                     values=("", f"{meta.row_count:,}"))
            self._cb_tree.item(fid, open=True)

    def _combine_refresh_cols(self):
        """Rebuild column checkboxes from union of all loaded files."""
        # Compute the new union of columns
        new_cols: list[str] = []
        for sheets in self._cb_files.values():
            for meta in sheets.values():
                for col in meta.columns:
                    if str(col) not in new_cols:
                        new_cols.append(str(col))
        # Skip rebuild if columns haven't changed
        if list(self._cb_col_vars.keys()) == new_cols:
            return
        for w in self._cb_col_scroll.winfo_children():
            w.destroy()
        self._cb_col_vars.clear()
        seen: dict[str, BooleanVar] = {}
        for col in new_cols:
            v = BooleanVar(value=True)
            ctk.CTkCheckBox(self._cb_col_scroll, text=col,
                            variable=v).pack(anchor="w", padx=4, pady=1)
            seen[col] = v
        self._cb_col_vars = seen
        self._cb_dedup_key_menu.configure(values=["— toutes les colonnes —"] + new_cols)
        if self._cb_dedup_key_var.get() not in self._cb_dedup_key_menu.cget("values"):
            self._cb_dedup_key_var.set("— toutes les colonnes —")

    def _combine_toggle_all(self, state: bool):
        for v in self._cb_col_vars.values():
            v.set(state)

    def _build_combine_panel(self, parent) -> ctk.CTkFrame:
        """Standalone panel to combine multiple Excel/CSV files."""
        f = ctk.CTkFrame(parent, fg_color="transparent")

        # ── File picker ───────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(f, fg_color="transparent")
        btn_row.pack(fill="x", pady=(4, 4))
        ctk.CTkButton(btn_row, text="Ajouter des fichiers …", width=170,
                      command=self._combine_browse).pack(side="left")
        ctk.CTkButton(btn_row, text="Retirer sélection", width=130,
                      fg_color="#c0392b", hover_color="#992d22",
                      command=self._combine_remove).pack(side="left", padx=8)

        list_fr = ctk.CTkFrame(f)
        list_fr.pack(fill="both", expand=True, pady=(0, 4))
        list_fr.columnconfigure(0, weight=1)
        list_fr.rowconfigure(0, weight=1)
        self._cb_tree = ttk.Treeview(
            list_fr, columns=("sheets", "rows"), height=6, show="tree headings")
        self._cb_tree.heading("#0",     text="Fichier")
        self._cb_tree.heading("sheets", text="Onglets")
        self._cb_tree.heading("rows",   text="Lignes")
        self._cb_tree.column("#0",     width=320, stretch=True)
        self._cb_tree.column("sheets", width=70, anchor="center")
        self._cb_tree.column("rows",   width=80, anchor="center")
        cb_vsb = ttk.Scrollbar(list_fr, orient="vertical", command=self._cb_tree.yview)
        self._cb_tree.configure(yscrollcommand=cb_vsb.set)
        self._cb_tree.grid(row=0, column=0, sticky="nsew")
        cb_vsb.grid(row=0, column=1, sticky="ns")
        self._cb_files: dict[str, dict[str, SheetMetadata]] = {}

        # ── Options ─────────────────────────────────────────────────────────
        opt_row = ctk.CTkFrame(f, fg_color="transparent")
        opt_row.pack(fill="x", pady=(4, 4))
        self._cb_dedup = BooleanVar(value=False)
        ctk.CTkCheckBox(opt_row, text="Supprimer les doublons",
                        variable=self._cb_dedup).pack(side="left", padx=(0, 20))
        ctk.CTkLabel(opt_row, text="Clé de dédoublon :",
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 6))
        self._cb_dedup_key_var = StringVar(value="— toutes les colonnes —")
        self._cb_dedup_key_menu = ctk.CTkOptionMenu(
            opt_row, variable=self._cb_dedup_key_var,
            values=["— toutes les colonnes —"], width=220)
        self._cb_dedup_key_menu.pack(side="left")

        # ── Column selector ───────────────────────────────────────────────
        col_frame = ctk.CTkFrame(f)
        col_frame.pack(fill="both", expand=True, pady=(4, 0))
        col_hdr = ctk.CTkFrame(col_frame, fg_color="transparent")
        col_hdr.pack(fill="x", padx=6, pady=(6, 4))
        ctk.CTkLabel(col_hdr, text="Colonnes à inclure dans le résultat :",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")
        ctk.CTkButton(col_hdr, text="Rien", width=50, height=22,
                      command=lambda: self._combine_toggle_all(False)).pack(side="right", padx=2)
        ctk.CTkButton(col_hdr, text="Tout", width=50, height=22,
                      command=lambda: self._combine_toggle_all(True)).pack(side="right", padx=2)
        self._cb_col_scroll = ctk.CTkScrollableFrame(col_frame, height=120)
        self._cb_col_scroll.pack(fill="both", expand=True, padx=4, pady=(0, 6))
        self._cb_col_vars: dict[str, BooleanVar] = {}

        return f

    def _build_result_section(self):
        sec = ctk.CTkFrame(self._scroll)
        sec.pack(fill="both", expand=True, padx=4, pady=6)
        ctk.CTkLabel(sec, text="3 · Résultat & Export",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     ).pack(anchor="w", padx=10, pady=(8, 4))

        # Stats
        self._stats_var = StringVar(value="Lancez une opération pour voir les résultats.")
        ctk.CTkLabel(sec, textvariable=self._stats_var,
                     font=ctk.CTkFont(size=12)).pack(anchor="w", padx=10)

        # Result Treeview
        tv_host = ctk.CTkFrame(sec, fg_color="transparent")
        tv_host.pack(fill="both", expand=True, padx=8, pady=(6, 4))
        tv_host.columnconfigure(0, weight=1)
        tv_host.rowconfigure(0, weight=1)

        self._rtree = ttk.Treeview(tv_host, height=12, show="headings")
        vsb = ttk.Scrollbar(tv_host, orient="vertical",   command=self._rtree.yview)
        hsb = ttk.Scrollbar(tv_host, orient="horizontal",  command=self._rtree.xview)
        self._rtree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._rtree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # ── "Non trouvés" panel (intersect only) ─────────────────────────
        self._nf_frame = ctk.CTkFrame(sec)
        # (not packed yet – shown only after an intersect)
        self._nf_stats_var = StringVar(value="")
        ctk.CTkLabel(self._nf_frame, textvariable=self._nf_stats_var,
                     font=ctk.CTkFont(size=12),
                     text_color="#e67e22").pack(anchor="w", padx=10, pady=(6, 2))

        nf_tv = ctk.CTkFrame(self._nf_frame, fg_color="transparent")
        nf_tv.pack(fill="both", expand=True, padx=8, pady=(4, 6))
        nf_tv.columnconfigure(0, weight=1)
        nf_tv.rowconfigure(0, weight=1)
        self._nf_tree = ttk.Treeview(nf_tv, height=8, show="headings")
        nf_vsb = ttk.Scrollbar(nf_tv, orient="vertical",  command=self._nf_tree.yview)
        nf_hsb = ttk.Scrollbar(nf_tv, orient="horizontal", command=self._nf_tree.xview)
        self._nf_tree.configure(yscrollcommand=nf_vsb.set, xscrollcommand=nf_hsb.set)
        self._nf_tree.grid(row=0, column=0, sticky="nsew")
        nf_vsb.grid(row=0, column=1, sticky="ns")
        nf_hsb.grid(row=1, column=0, sticky="ew")

        # Columns to keep
        ck_host = ctk.CTkFrame(sec)
        ck_host.pack(fill="x", padx=8, pady=(4, 4))
        ctk.CTkLabel(ck_host, text="Colonnes à exporter :",
                     font=ctk.CTkFont(size=12)).pack(anchor="w", padx=8, pady=(6, 2))
        self._col_scroll = ctk.CTkScrollableFrame(ck_host, height=70,
                                                    orientation="horizontal")
        self._col_scroll.pack(fill="x", padx=8, pady=(0, 6))

        # Export row
        exp = ctk.CTkFrame(sec, fg_color="transparent")
        exp.pack(fill="x", padx=8, pady=(4, 8))
        ctk.CTkLabel(exp, text="Fichier de sortie :").pack(side="left", padx=(0, 6))
        self._out_var = StringVar()
        ctk.CTkEntry(exp, textvariable=self._out_var,
                      width=320, placeholder_text="(choisissez le chemin de sauvegarde …)").pack(side="left")
        ctk.CTkButton(exp, text="…", width=36,
                       command=self._pick_output).pack(side="left", padx=4)
        ctk.CTkButton(exp, text="Enregistrer en Excel",
                       command=lambda: self._export("xlsx")).pack(side="left", padx=(14, 4))
        ctk.CTkButton(exp, text="Enregistrer en CSV",
                       command=lambda: self._export("csv")).pack(side="left", padx=4)

        # History log
        log_host = ctk.CTkFrame(sec)
        log_host.pack(fill="x", padx=8, pady=(0, 8))
        ctk.CTkLabel(log_host, text="Historique",
                     font=ctk.CTkFont(size=11, weight="bold")).pack(
            anchor="w", padx=8, pady=(4, 0))
        self._log_box = ctk.CTkTextbox(
            log_host, height=80,
            font=ctk.CTkFont(size=11, family="Courier New"))
        self._log_box.pack(fill="x", padx=8, pady=(0, 6))
        self._log_box.configure(state="disabled")

    # =========================================================================
    # Event callbacks
    # =========================================================================

    def _file_loaded(self):
        a_cols = list(self._pa.df.columns) if self._pa.df is not None else []
        b_cols = list(self._pb.df.columns) if self._pb.df is not None else []

        if a_cols:
            self._mi_a_menu.configure(values=a_cols)
            if self._mi_a_var.get() not in a_cols:
                self._mi_a_var.set(a_cols[0])
            self._ex_col_menu.configure(values=a_cols)
            if self._ex_col_var.get() not in a_cols:
                self._ex_col_var.set(a_cols[0])
            self._en_a_menu.configure(values=a_cols)
            if self._en_a_var.get() not in a_cols:
                self._en_a_var.set(a_cols[0])

        if b_cols:
            self._mi_b_menu.configure(values=b_cols)
            if self._mi_b_var.get() not in b_cols:
                self._mi_b_var.set(b_cols[0])
            self._en_b_menu.configure(values=b_cols)
            if self._en_b_var.get() not in b_cols:
                self._en_b_var.set(b_cols[0])

        all_cols = a_cols + [c for c in b_cols if c not in a_cols]
        if all_cols:
            self._ext_key_menu.configure(
                values=["— aucune (toutes les colonnes) —"] + all_cols)
        self._refresh_mi_cols()
        self._refresh_en_cols()

    def _op_changed(self, val: str):
        self._op_desc_var.set(OP_DESC.get(val, ""))
        if "Extraire" in val:
            key = "extract"
        elif "Fusionner" in val:
            key = "extend"
        elif "Enrichir" in val:
            key = "enrich"
        elif "Combiner" in val:
            key = "combine"
        elif val == "Scinder  (Onglets)":
            key = "split"
        elif val == "Scinder  (Colonnes)":
            key = "splitcol"
        elif val == "Scinder  (Lignes)":
            key = "splitlines"
        else:
            key = "mi"
        self._show_sub(key)
        # Show column selector only for intersect
        if "Intersecter" in val:
            self._mi_col_frame.grid()
        else:
            self._mi_col_frame.grid_remove()
        # Hide file panels for standalone operations
        is_standalone = val in ("Scinder  (Onglets)", "Combiner  (Multi-fichiers)",
                                "Scinder  (Colonnes)", "Scinder  (Lignes)")
        if is_standalone:
            self._files_sec.pack_forget()
        else:
            siblings = self._scroll.winfo_children()
            if self._files_sec not in [w for w in siblings if w.winfo_ismapped()]:
                self._files_sec.pack(fill="x", padx=4, pady=6,
                                      before=self._op_sec)

    def _cond_changed(self, val: str):
        if val in ("est vide", "n'est pas vide"):
            self._ex_val_entry.configure(state="disabled")
        else:
            self._ex_val_entry.configure(state="normal")

    def _en_fill_changed(self, val: str):
        if val == "Valeur personnalisée":
            self._en_fill_entry.configure(state="normal")
        else:
            self._en_fill_entry.configure(state="disabled")

    def _show_sub(self, key: str):
        for p in self._sub.values():
            p.pack_forget()
        self._sub[key].pack(fill="x")

    def _refresh_mi_cols(self):
        """Rebuild the column checkboxes in the intersect column selector."""
        mode = self._mi_mode_var.get()

        # Resolve A columns
        if mode == "Plusieurs A × 1 B":
            a_cols: list[str] = []
            for df in self._mi_multi_files.values():
                for c in df.columns:
                    if str(c) not in a_cols:
                        a_cols.append(str(c))
        else:
            a_cols = list(self._pa.df.columns) if self._pa.df is not None else []

        # Resolve B columns
        if mode == "1 A × Plusieurs B":
            b_cols: list[str] = []
            for df in self._mi_multi_files.values():
                for c in df.columns:
                    if str(c) not in b_cols:
                        b_cols.append(str(c))
        else:
            b_cols = list(self._pb.df.columns) if self._pb.df is not None else []

        # Skip rebuild if columns haven't changed
        if (list(self._mi_a_col_vars.keys()) == a_cols
                and list(self._mi_b_col_vars.keys()) == b_cols):
            return
        for w in self._mi_a_col_scroll.winfo_children():
            w.destroy()
        for w in self._mi_b_col_scroll.winfo_children():
            w.destroy()
        self._mi_a_col_vars.clear()
        self._mi_b_col_vars.clear()
        if a_cols:
            for col in a_cols:
                v = BooleanVar(value=True)
                ctk.CTkCheckBox(self._mi_a_col_scroll, text=str(col),
                                variable=v).pack(anchor="w", padx=4, pady=1)
                self._mi_a_col_vars[col] = v
        if b_cols:
            for col in b_cols:
                v = BooleanVar(value=False)
                ctk.CTkCheckBox(self._mi_b_col_scroll, text=str(col),
                                variable=v).pack(anchor="w", padx=4, pady=1)
                self._mi_b_col_vars[col] = v

    def _mi_toggle_all(self, which: str, state: bool):
        """Toggle all checkboxes in A or B column list."""
        vars_ = self._mi_a_col_vars if which == "a" else self._mi_b_col_vars
        for v in vars_.values():
            v.set(state)

    # ── Multi-file mode helpers ──────────────────────────────────────────────

    def _mi_mode_changed(self, val: str):
        """Show/hide file panels and multi-file picker based on mode."""
        if val == "Standard (1 A × 1 B)":
            self._mi_multi_frame.grid_remove()
            # Ensure both file panels are visible
            self._pa.pack(fill="x", padx=6, pady=(0, 4))
            self._pb.pack(fill="x", padx=6, pady=(0, 8))
        elif val == "1 A × Plusieurs B":
            # Keep A panel visible, hide B, show multi picker for B
            self._pa.pack(fill="x", padx=6, pady=(0, 4))
            self._pb.pack_forget()
            self._mi_multi_label.configure(text="Fichiers B (plusieurs) :")
            self._mi_multi_frame.grid()
        elif val == "Plusieurs A × 1 B":
            # Keep B panel visible, hide A, show multi picker for A
            self._pa.pack_forget()
            self._pb.pack(fill="x", padx=6, pady=(0, 8))
            self._mi_multi_label.configure(text="Fichiers A (plusieurs) :")
            self._mi_multi_frame.grid()
        # Force-refresh the column checkboxes for the new mode
        self._mi_a_col_vars.clear()
        self._mi_b_col_vars.clear()
        self._refresh_mi_cols()

    def _mi_multi_browse(self):
        """Add files to the multi-file list."""
        paths = filedialog.askopenfilenames(
            title="Sélectionner des fichiers",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Tous", "*.*")],
        )
        if not paths:
            return
        for p in paths:
            if p in self._mi_multi_files:
                continue
            try:
                df = load_file(p)
                self._mi_multi_files[p] = df
            except Exception as e:
                messagebox.showerror("Erreur", f"{pathlib.Path(p).name}\n{e}")
        self._mi_multi_refresh_tree()
        self._mi_multi_refresh_key()
        self._refresh_mi_cols()

    def _mi_multi_remove(self):
        """Remove selected files from the multi-file list."""
        sel = self._mi_multi_tree.selection()
        if not sel:
            return
        for iid in sel:
            # If a child (column row) is selected, resolve to the parent file
            parent = self._mi_multi_tree.parent(iid)
            path = iid if not parent else parent
            self._mi_multi_files.pop(path, None)
        self._mi_multi_refresh_tree()
        self._mi_multi_refresh_key()
        self._refresh_mi_cols()

    def _mi_multi_refresh_tree(self):
        self._mi_multi_tree.delete(*self._mi_multi_tree.get_children())
        for path, df in self._mi_multi_files.items():
            fid = self._mi_multi_tree.insert(
                "", "end", iid=path,
                values=(len(df.columns), f"{len(df):,}"),
                text=pathlib.Path(path).name,
            )
            # Show column names as children
            for col in df.columns:
                self._mi_multi_tree.insert(
                    fid, "end",
                    text=f"  ↳ {col}",
                    values=("", ""),
                )
            self._mi_multi_tree.item(fid, open=True)
        # Switch to tree+headings so file names are visible
        self._mi_multi_tree.configure(show="tree headings")
        self._mi_multi_tree.heading("#0", text="Fichier / Colonnes")
        self._mi_multi_tree.column("#0", width=260, stretch=True)

    def _mi_multi_refresh_key(self):
        """Refresh the key-column dropdown from union of multi-file columns."""
        all_cols: list[str] = []
        for df in self._mi_multi_files.values():
            for c in df.columns:
                if str(c) not in all_cols:
                    all_cols.append(str(c))
        if all_cols:
            self._mi_multi_key_menu.configure(values=all_cols)
            if self._mi_multi_key_var.get() not in all_cols:
                self._mi_multi_key_var.set(all_cols[0])
        else:
            self._mi_multi_key_menu.configure(values=["— ajouter des fichiers —"])
            self._mi_multi_key_var.set("— ajouter des fichiers —")

    def _refresh_en_cols(self):
        """Rebuild the B-column checkboxes in the Enrichir panel."""
        b_cols = list(self._pb.df.columns) if self._pb.df is not None else []
        # Skip rebuild if columns haven't changed
        if list(self._en_col_vars.keys()) == b_cols:
            return
        for w in self._en_col_scroll.winfo_children():
            w.destroy()
        self._en_col_vars.clear()
        if b_cols:
            for col in b_cols:
                v = BooleanVar(value=True)
                ctk.CTkCheckBox(self._en_col_scroll, text=str(col),
                                variable=v).pack(anchor="w", padx=4, pady=1)
                self._en_col_vars[col] = v

    def _en_toggle_all(self, state: bool):
        """Toggle all checkboxes in the Enrichir B-column list."""
        for v in self._en_col_vars.values():
            v.set(state)

    def _toggle_theme(self):
        if self._theme_sw.get():
            ctk.set_appearance_mode("light")
            self._theme_sw.configure(text="Mode sombre")
        else:
            ctk.set_appearance_mode("dark")
            self._theme_sw.configure(text="Mode clair")

    def _clear_all(self):
        for panel in (self._pa, self._pb):
            panel.df   = None
            panel.path = ""
            panel._fname_var.set("Aucun fichier sélectionné")
            panel._info_var.set("")
            panel._tree.delete(*panel._tree.get_children())
            panel._tree["columns"] = []
        self._result_df = None
        self._notfound_df = None
        self._en_found_df = None
        self._rtree.delete(*self._rtree.get_children())
        self._rtree["columns"] = []
        self._hide_notfound()
        for w in self._col_scroll.winfo_children():
            w.destroy()
        self._col_vars.clear()
        self._refresh_mi_cols()
        self._refresh_en_cols()
        self._mi_multi_files.clear()
        self._mi_multi_refresh_tree()
        self._mi_multi_refresh_key()
        self._mi_mode_var.set("Standard (1 A × 1 B)")
        self._mi_multi_frame.grid_remove()
        self._sp_files.clear()
        self._split_refresh_tree()
        self._cb_files.clear()
        self._combine_refresh_tree()
        self._combine_refresh_cols()
        # Reset Scinder (Colonnes)
        self._sc_df = None
        self._sc_fname_var.set("Aucun fichier sélectionné")
        self._sc_split_cols.clear()
        self._sc_split_list.delete(*self._sc_split_list.get_children())
        self._sc_rebuild_cols()
        self._sc_dir_var.set("")

        if hasattr(self, "_sl_fname_var"):
            self._sl_df = None
            self._sl_fname_var.set("Aucun fichier sélectionné")
            self._sl_num_var.set("200")
            self._sl_dir_var.set("")

        self._stats_var.set("Lancez une opération pour voir les résultats.")
        self._out_var.set("")
        self._log("État réinitialisé.")

    # =========================================================================
    # Operations
    # =========================================================================

    def _run(self):
        op = self._op_var.get()
        self._notfound_df = None
        self._hide_notfound()
        self._stats_var.set("Opération en cours …")
        self.update_idletasks()

        def _compute():
            try:
                if "Soustraire" in op:
                    result = self._op_minus_or_intersect(minus=True)
                elif "Fusionner" in op:
                    result = self._op_extend()
                elif "Intersecter" in op:
                    result = self._op_minus_or_intersect(minus=False)
                elif "Enrichir" in op:
                    result = self._op_enrich()
                elif "Extraire" in op:
                    result = self._op_extract()
                elif "Combiner" in op:
                    result = self._op_combine()
                elif op == "Scinder  (Colonnes)":
                    self.after(0, self._op_splitcol)
                    return
                elif op == "Scinder  (Lignes)":
                    self.after(0, self._op_splitlines)
                    return
                elif op == "Scinder  (Onglets)":
                    self.after(0, self._op_split)
                    return
                else:
                    raise ValueError("Opération inconnue")
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Erreur d'opération", str(e)))
                self.after(0, lambda: self._stats_var.set(
                    "Lancez une opération pour voir les résultats."))
                return

            def _show_result():
                self._result_df = result
                a_n = len(self._pa.df) if self._pa.df is not None else None
                b_n = len(self._pb.df) if self._pb.df is not None else None
                nf_n = len(self._notfound_df) if self._notfound_df is not None else None
                ef_n = len(self._en_found_df) if self._en_found_df is not None else None
                stats = (
                    f"Fichier A : {fmt(a_n)} lignes   |   "
                    f"Fichier B : {fmt(b_n)} lignes   |   "
                    f"Résultat : {len(result):,} lignes"
                )
                if nf_n is not None:
                    stats += f"   |   Non trouvés : {nf_n:,} lignes"
                if ef_n is not None:
                    stats += f"   |   Trouvés B : {ef_n:,} lignes"
                self._stats_var.set(stats)
                self._fill_result_tree(result)
                self._refresh_col_checks(result)
                if self._notfound_df is not None and len(self._notfound_df):
                    self._show_notfound(self._notfound_df)
                self._log(
                    f"[{op.strip()}]  A={fmt(a_n)}  B={fmt(b_n)}  → {len(result):,} lignes"
                    + (f"  | non trouvés={nf_n:,}" if nf_n else "")
                    + (f"  | trouvés B={ef_n:,}" if ef_n else "")
                )
            self.after(0, _show_result)

        threading.Thread(target=_compute, daemon=True).start()

    def _op_minus_or_intersect(self, minus: bool) -> pd.DataFrame:
        mode = self._mi_mode_var.get()

        # ── Resolve A side (single or multi) ────────────────────────────
        if mode == "Plusieurs A × 1 B":
            if not self._mi_multi_files:
                raise ValueError("Ajoutez au moins un Fichier A.")
            if self._pb.df is None:
                raise ValueError("Le Fichier B n'est pas chargé.")
            col_b = self._mi_b_var.get()
            if col_b not in self._pb.df.columns:
                raise ValueError(f"Colonne '{col_b}' introuvable dans le Fichier B.")
            multi_key = self._mi_multi_key_var.get()
            set_b = self._processor.build_membership_set(self._pb.df, col_b)
            frames: list[pd.DataFrame] = []
            all_a_keys: set[str] = set()
            for path, df in self._mi_multi_files.items():
                if multi_key not in df.columns:
                    raise ValueError(
                        f"Colonne '{multi_key}' introuvable dans "
                        f"{pathlib.Path(path).name}.")
                filtered, key_set = self._processor.filter_by_membership(df, multi_key, set_b, minus=minus)
                frames.append(filtered)
                all_a_keys.update(key_set)
            result = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

            # Not-found: B keys absent from ALL A files
            if not minus and self._mi_nf_var.get():
                nf = self._processor.rows_not_in_membership(self._pb.df, col_b, all_a_keys)
                if len(nf):
                    self._notfound_df = nf
            return result

        # ── Resolve B side (single or multi) ────────────────────────────
        if mode == "1 A × Plusieurs B":
            if self._pa.df is None:
                raise ValueError("Le Fichier A n'est pas chargé.")
            if not self._mi_multi_files:
                raise ValueError("Ajoutez au moins un Fichier B.")
            col_a = self._mi_a_var.get()
            if col_a not in self._pa.df.columns:
                raise ValueError(f"Colonne '{col_a}' introuvable dans le Fichier A.")
            multi_key = self._mi_multi_key_var.get()
            # Union all B keys
            set_b: set[str] = set()
            for path, df in self._mi_multi_files.items():
                if multi_key not in df.columns:
                    raise ValueError(
                        f"Colonne '{multi_key}' introuvable dans "
                        f"{pathlib.Path(path).name}.")
                set_b.update(self._processor.build_membership_set(df, multi_key))
            result, set_a = self._processor.filter_by_membership(self._pa.df, col_a, set_b, minus=minus)

            # Not-found: union of B keys absent from A
            if not minus and self._mi_nf_var.get():
                nf_frames: list[pd.DataFrame] = []
                for df in self._mi_multi_files.values():
                    nf_frames.append(self._processor.rows_not_in_membership(df, multi_key, set_a))
                if nf_frames:
                    nf = pd.concat(nf_frames, ignore_index=True)
                    if len(nf):
                        self._notfound_df = nf
            return result

        # ── Standard (1 A × 1 B) ────────────────────────────────────────
        if self._pa.df is None:
            raise ValueError("Le Fichier A n'est pas chargé.")
        if self._pb.df is None:
            raise ValueError("Le Fichier B n'est pas chargé.")
        col_a = self._mi_a_var.get()
        col_b = self._mi_b_var.get()
        if col_a not in self._pa.df.columns:
            raise ValueError(f"Colonne '{col_a}' introuvable dans le Fichier A.")
        if col_b not in self._pb.df.columns:
            raise ValueError(f"Colonne '{col_b}' introuvable dans le Fichier B.")

        if minus:
            return self._processor.subtract(self._pa.df, col_a, self._pb.df, col_b)

        op_result = self._processor.intersect(
            self._pa.df,
            col_a,
            self._pb.df,
            col_b,
            selected_a=[c for c, v in self._mi_a_col_vars.items() if v.get()],
            selected_b=[c for c, v in self._mi_b_col_vars.items() if v.get()],
            principal=self._mi_principal_var.get(),
            suffix=self._mi_suffix_var.get() or " (B)",
            include_not_found=self._mi_nf_var.get(),
        )
        self._notfound_df = op_result.not_found
        return op_result.dataframe

    def _op_extend(self) -> pd.DataFrame:
        if self._pa.df is None:
            raise ValueError("Le Fichier A n'est pas chargé.")
        if self._pb.df is None:
            raise ValueError("Le Fichier B n'est pas chargé.")
        return self._processor.extend(
            self._pa.df,
            self._pb.df,
            dedup=self._ext_dedup.get(),
            dedup_key=self._ext_key_var.get(),
        )

    def _op_enrich(self) -> pd.DataFrame:
        """Enrichir (A ← B): LEFT or INNER JOIN on key columns, bringing selected B columns."""
        if self._pa.df is None:
            raise ValueError("Le Fichier A n'est pas chargé.")
        if self._pb.df is None:
            raise ValueError("Le Fichier B n'est pas chargé.")

        col_a = self._en_a_var.get()
        col_b = self._en_b_var.get()
        if col_a not in self._pa.df.columns:
            raise ValueError(f"Colonne '{col_a}' introuvable dans le Fichier A.")
        if col_b not in self._pb.df.columns:
            raise ValueError(f"Colonne '{col_b}' introuvable dans le Fichier B.")

        sel_b = [c for c, v in self._en_col_vars.items() if v.get()]
        op_result = self._processor.enrich(
            self._pa.df,
            col_a,
            self._pb.df,
            col_b,
            selected_b=sel_b,
            suffix=self._en_suffix_var.get() or " (B)",
            join_mode=self._en_mode_var.get(),
            skip_empty=self._en_skip_empty.get(),
            fill_mode=self._en_fill_var.get(),
            fill_value=self._en_fill_entry.get(),
        )
        self._en_found_df = op_result.found
        return op_result.dataframe

    def _op_split(self):
        """Split every sheet of every selected file into individual files."""
        if not self._sp_files:
            messagebox.showwarning("Aucun fichier",
                                  "Ajoutez au moins un fichier Excel à scinder.")
            return
        pattern = self._sp_name_var.get().strip() or "{fichier}_{onglet}"
        out_fmt = self._sp_fmt_var.get()
        out_dir = self._sp_dir_var.get().strip()
        try:
            total_written = split_workbook_sheets(
                self._sp_files.keys(),
                pattern=pattern,
                output_format=out_fmt,
                output_dir=out_dir or None,
            )
            summary = (
                f"{total_written} fichier(s) créé(s) à partir de "
                f"{len(self._sp_files)} source(s)."
            )
            self._stats_var.set(summary)
            self._log(f"[Scinder]  {summary}")
            messagebox.showinfo("Scission terminée", summary)
        except Exception as e:
            messagebox.showerror("Erreur de scission", str(e))

    def _op_splitcol(self):
        """Split file(s) by column values in a background worker."""
        if self._sc_df is None:
            messagebox.showwarning("Aucun fichier", "Chargez un fichier dans Scinder (Colonnes).")
            return
        if not self._sc_split_cols:
            messagebox.showwarning("Colonnes manquantes",
                                   "Ajoutez au moins une colonne de découpe.")
            return
        for col in self._sc_split_cols:
            if col not in self._sc_df.columns:
                messagebox.showerror("Colonne introuvable",
                                     f"La colonne '{col}' n'existe pas dans le fichier chargé.")
                return

        keep_cols = [c for c, v in self._sc_keep_vars.items() if v.get()]
        if not keep_cols:
            messagebox.showwarning("Aucune colonne",
                                   "Cochez au moins une colonne à conserver.")
            return

        out_dir_str = self._sc_dir_var.get().strip()
        src_path = getattr(self, "_sc_source_path", "") or ""
        if out_dir_str:
            out_dir = pathlib.Path(out_dir_str)
        else:
            out_dir = pathlib.Path(src_path).parent if src_path else pathlib.Path(".")

        zebra = self._sc_zebra_var.get()
        split_cols = list(self._sc_split_cols)
        all_sheets_mode = (
            getattr(self, "_sc_all_sheets_var", None) is not None
            and self._sc_all_sheets_var.get()
            and self._sc_xls is not None
        )
        sheet_names = list(self._sc_xls.sheet_names) if all_sheets_mode else None
        single_sheet_name = self._sc_sheet_var.get() if self._sc_xls is not None else None
        self._sc_prog_frame.pack(fill="x", padx=4, pady=(2, 4))
        self._sc_prog_bar.set(0)
        self._sc_prog_label.configure(text="Préparation …")
        self.update_idletasks()

        def _progress(done: int, total: int, message: str):
            self.after(0, lambda: (
                self._sc_prog_bar.set(done / total if total else 0),
                self._sc_prog_label.configure(text=message),
            ))

        def _worker():
            total_written = 0
            error_msg = None
            try:
                total_written = split_by_columns(
                    source_path=src_path,
                    dataframe=self._sc_df,
                    split_columns=split_cols,
                    keep_columns=keep_cols,
                    output_dir=str(out_dir),
                    zebra=zebra,
                    all_sheets=all_sheets_mode,
                    sheet_names=sheet_names,
                    selected_sheet=single_sheet_name,
                    progress_callback=_progress,
                )
            except Exception as e:
                error_msg = str(e)

            # ── Finish on UI thread ─────────────────────────────────────
            def _done():
                if error_msg:
                    self._sc_prog_label.configure(text=f"Erreur : {error_msg}")
                    messagebox.showerror("Erreur de scission (colonnes)", error_msg)
                else:
                    self._sc_prog_bar.set(1)
                    summary = f"{total_written} fichier(s) créé(s) dans {out_dir}"
                    self._sc_prog_label.configure(text=f"✓ {summary}")
                    self._stats_var.set(summary)
                    self._log(f"[Scinder Colonnes]  {total_written} fichier(s) → {out_dir}")
                    messagebox.showinfo("Scission terminée", summary)
            self.after(0, _done)

        threading.Thread(target=_worker, daemon=True).start()


    def _op_extract(self) -> pd.DataFrame:
        if self._pa.df is None:
            raise ValueError("Le Fichier A n'est pas chargé.")
        col  = self._ex_col_var.get()
        if col not in self._pa.df.columns:
            raise ValueError(f"Colonne '{col}' introuvable dans le Fichier A.")
        return self._processor.extract(
            self._pa.df,
            col,
            self._ex_cond_var.get(),
            self._ex_val_var.get(),
            case_sensitive=self._ex_case_var.get(),
        )

    def _op_combine(self) -> pd.DataFrame:
        """Combiner: stack multiple files vertically, keeping selected columns."""
        if not self._cb_files:
            raise ValueError("Ajoutez au moins un fichier à combiner.")

        sel_cols = [c for c, v in self._cb_col_vars.items() if v.get()]

        frames: list[pd.DataFrame] = []
        for path, sheets in self._cb_files.items():
            is_excel = pathlib.Path(path).suffix.lower() in EXCEL_SUFFIXES
            for sheet_name in sheets:
                sheet_arg = sheet_name if is_excel else 0
                frames.append(load_tabular_file(path, sheet_name=sheet_arg))

        result = self._processor.combine(
            frames,
            selected_columns=sel_cols,
            dedup=self._cb_dedup.get(),
            dedup_key=self._cb_dedup_key_var.get(),
        )

        self._cb_dedup_key_menu.configure(
            values=["— toutes les colonnes —"] + list(result.columns))

        return result

    # =========================================================================
    # Result table helpers
    # =========================================================================

    def _fill_result_tree(self, df: pd.DataFrame, max_rows: int = 100):
        self._rtree.delete(*self._rtree.get_children())
        self._rtree["columns"] = list(df.columns)
        for c in df.columns:
            self._rtree.heading(c, text=c)
            self._rtree.column(c, width=110, minwidth=60, stretch=False)
        for row in df.head(max_rows).values.tolist():
            self._rtree.insert("", "end", values=row)

    def _refresh_col_checks(self, df: pd.DataFrame):
        for w in self._col_scroll.winfo_children():
            w.destroy()
        self._col_vars.clear()
        for col in df.columns:
            v = BooleanVar(value=True)
            ctk.CTkCheckBox(self._col_scroll, text=str(col),
                             variable=v).pack(side="left", padx=6)
            self._col_vars[col] = v

    # =========================================================================
    # Not-found panel helpers
    # =========================================================================

    def _show_notfound(self, df: pd.DataFrame):
        self._nf_stats_var.set(
            f"⚠  Non trouvés dans Fichier A : {len(df):,} lignes du Fichier B"
        )
        self._nf_tree.delete(*self._nf_tree.get_children())
        self._nf_tree["columns"] = list(df.columns)
        for c in df.columns:
            self._nf_tree.heading(c, text=c)
            self._nf_tree.column(c, width=110, minwidth=60, stretch=False)
        for row in df.head(100).values.tolist():
            self._nf_tree.insert("", "end", values=row)
        self._nf_frame.pack(fill="both", expand=True, padx=8, pady=(4, 6))

    def _hide_notfound(self):
        self._nf_frame.pack_forget()
        self._nf_tree.delete(*self._nf_tree.get_children())
        self._nf_tree["columns"] = []
        self._nf_stats_var.set("")

    # =========================================================================
    # Export
    # =========================================================================

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Enregistrer le résultat sous",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("Tous", "*.*")],
        )
        if path:
            self._out_var.set(path)

    def _export(self, fmt_: str):
        if self._result_df is None:
            messagebox.showwarning("Aucun résultat", "Lancez d'abord une opération.")
            return
        path = self._out_var.get()
        if not path:
            path = filedialog.asksaveasfilename(
                title="Enregistrer le résultat",
                defaultextension=f".{fmt_}",
                filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("Tous", "*.*")],
            )
            if not path:
                return
            self._out_var.set(path)

        selected = [c for c, v in self._col_vars.items() if v.get()]
        if not selected:
            selected = list(self._result_df.columns)
        df_out = self._result_df[selected]

        has_extra = (
            (self._notfound_df is not None and len(self._notfound_df))
            or (self._en_found_df is not None and len(self._en_found_df))
        )

        try:
            if path.endswith(".csv"):
                df_out.to_csv(path, index=False, encoding="utf-8-sig")
                # Save not-found as a separate CSV if present
                if self._notfound_df is not None and len(self._notfound_df):
                    nf_path = path.replace(".csv", "_non-trouvés.csv")
                    self._notfound_df.to_csv(nf_path, index=False, encoding="utf-8-sig")
                    self._log(f"Non trouvés : {len(self._notfound_df):,} lignes  →  {pathlib.Path(nf_path).name}")
                if self._en_found_df is not None and len(self._en_found_df):
                    ef_path = path.replace(".csv", "_trouvés-B.csv")
                    self._en_found_df.to_csv(ef_path, index=False, encoding="utf-8-sig")
                    self._log(f"Trouvés B : {len(self._en_found_df):,} lignes  →  {pathlib.Path(ef_path).name}")
            else:
                if has_extra:
                    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
                        df_out.to_excel(writer, sheet_name="Résultat", index=False)
                        if self._notfound_df is not None and len(self._notfound_df):
                            self._notfound_df.to_excel(writer, sheet_name="Non trouvés", index=False)
                        if self._en_found_df is not None and len(self._en_found_df):
                            self._en_found_df.to_excel(writer, sheet_name="Trouvés dans B", index=False)
                else:
                    df_out.to_excel(path, index=False, engine="xlsxwriter")
            self._log(
                f"Enregistré : {len(df_out):,} lignes × {len(selected)} colonnes  →  "
                f"{pathlib.Path(path).name}"
            )
            messagebox.showinfo("Enregistré", f"Fichier sauvegardé :\n{path}")
        except Exception as e:
            messagebox.showerror("Erreur d'enregistrement", str(e))

    # =========================================================================
    # Profile save / load
    # =========================================================================

    def _profile_state(self) -> dict:
        return {
            "file_a":            self._pa.path,
            "file_b":            self._pb.path,
            "col_a":             self._mi_a_var.get(),
            "col_b":             self._mi_b_var.get(),
            "operation":         self._op_var.get(),
            "extract_col":       self._ex_col_var.get(),
            "extract_condition": self._ex_cond_var.get(),
            "extract_value":     self._ex_val_var.get(),
            "extract_case":      self._ex_case_var.get(),
            "extend_dedup":      self._ext_dedup.get(),
            "extend_key":        self._ext_key_var.get(),
            "output_path":       self._out_var.get(),
        }

    def _save_profile(self):
        path = filedialog.asksaveasfilename(
            title="Sauvegarder le profil",
            initialfile="filter_profile.json",
            initialdir=str(WORKSPACE),
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(self._profile_state(), fh, indent=2, ensure_ascii=False)
        self._log(f"Profil sauvegardé → {pathlib.Path(path).name}")

    def _load_profile(self):
        path = filedialog.askopenfilename(
            title="Charger un profil",
            initialdir=str(WORKSPACE),
            filetypes=[("JSON", "*.json")],
        )
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as fh:
                p = json.load(fh)
        except Exception as e:
            messagebox.showerror("Erreur de profil", str(e))
            return

        # Re-load data files
        for panel, key in [(self._pa, "file_a"), (self._pb, "file_b")]:
            fp = p.get(key, "")
            if fp and pathlib.Path(fp).exists():
                try:
                    panel._accept(fp, load_file(fp))
                except Exception:
                    pass
        self._file_loaded()

        # Restore operation
        if p.get("operation") in OPERATIONS:
            self._op_var.set(p["operation"])
            self._op_changed(p["operation"])

        def _set(menu: ctk.CTkOptionMenu, var: StringVar, val: str):
            if val and val in menu.cget("values"):
                var.set(val)

        _set(self._mi_a_menu,  self._mi_a_var,   p.get("col_a", ""))
        _set(self._mi_b_menu,  self._mi_b_var,   p.get("col_b", ""))
        _set(self._ex_col_menu, self._ex_col_var, p.get("extract_col", ""))
        if p.get("extract_condition") in CONDITIONS:
            self._ex_cond_var.set(p["extract_condition"])
        self._ex_val_var.set(p.get("extract_value", ""))
        self._ex_case_var.set(bool(p.get("extract_case", False)))
        self._ext_dedup.set(bool(p.get("extend_dedup", True)))
        _set(self._ext_key_menu, self._ext_key_var, p.get("extend_key", ""))
        self._out_var.set(p.get("output_path", ""))
        self._log(f"Profil chargé ← {pathlib.Path(path).name}")

    # =========================================================================
    # History log
    # =========================================================================

    def _log(self, msg: str):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self._log_box.configure(state="normal")
        self._log_box.insert("end", f"[{ts}]  {msg}\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        app = App()
        app.mainloop()
    except Exception as _e:
        import traceback
        from tkinter import messagebox as _mb
        import tkinter as _tk
        _r = _tk.Tk(); _r.withdraw()
        _mb.showerror("Erreur au démarrage", traceback.format_exc())
        _r.destroy()
