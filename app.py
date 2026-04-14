"""
Excel Filter Tool  — Generic GUI
Replaces the hardcoded filter.py with a fully interactive desktop app.
"""

import datetime
import logging
import os
import pathlib
import threading
from tkinter import filedialog, messagebox
from tkinter import ttk

import customtkinter as ctk
import polars as pl
# Enable string cache globally for improved categorical performance
pl.enable_string_cache()
import fastexcel
from tkinter import BooleanVar, StringVar
import pdf_generator
import operations
from export_utils import ensure_excel_output_path

# ── Defaults ──────────────────────────────────────────────────────────────────

from logging.handlers import RotatingFileHandler
logging.basicConfig(
    handlers=[RotatingFileHandler("error.log", maxBytes=1_000_000, backupCount=1, encoding="utf-8")],
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

WORKSPACE     = pathlib.Path(__file__).parent
DEFAULT_OUT   = pathlib.Path.home() / "Desktop" / "Résultats"
OPERATIONS    = [
    "Soustraire  (A − B)", "Fusionner  (A ∪ B)", "Intersecter  (A ∩ B)", 
    "Enrichir  (A ← B)", "Combiner  (Multi-fichiers)", "Extraire  (Filtrer A)", 
    "Split  (Sheets)", "Split  (Columns)", "Excel to PDF"
]
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
    "Combiner  (Multi-fichiers)": "Empile plusieurs fichiers Excel, choisissez les colonnes à conserver et dédoublonnez si souhaité.",
    "Extraire  (Filtrer A)": "Filtre les lignes de A selon une condition sur une colonne choisie.",
    "Split  (Sheets)":    "Sépare chaque onglet d'un ou plusieurs fichiers Excel en fichiers individuels.",
    "Split  (Columns)":   "Charge un seul fichier, choisissez les colonnes à conserver et une ou plusieurs colonnes de découpe. Génère un fichier Excel par combinaison de valeurs unique.",
    "Excel to PDF": "Génère des rapports PDF stylisés (en masse) à partir d'un dossier de fichiers Excel.",
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def load_file(path: str) -> pl.DataFrame:
    p = pathlib.Path(path)
    if p.suffix.lower() in (".xlsx", ".xls"):
        return operations.cleanup_floats(pl.read_excel(path))
    raise ValueError(f"Unsupported file type: {p.suffix}")


def fmt(n) -> str:
    """Format a row-count or return '—' for None."""
    return f"{n:,}" if isinstance(n, int) else "—"


def _show_done_dialog(parent, title: str, message: str, open_path: str):
    """Show a completion dialog with 'Ouvrir' and 'Fermer' buttons."""
    w, h = 460, 150
    dlg = ctk.CTkToplevel(parent)
    dlg.title("Opération Terminée")
    # Center on screen
    sx = dlg.winfo_screenwidth()
    sy = dlg.winfo_screenheight()
    dlg.geometry(f"{w}x{h}+{(sx - w) // 2}+{(sy - h) // 2}")
    dlg.resizable(False, False)
    dlg.grab_set()
    dlg.transient(parent)
    ctk.CTkLabel(dlg, text=message, wraplength=420,
                 font=ctk.CTkFont(size=12)).pack(padx=20, pady=(20, 14))
    btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
    btn_row.pack(pady=(0, 16))
    def _open():
        p = pathlib.Path(open_path)
        target = p if p.is_dir() else p.parent
        try:
            os.startfile(str(target))
        except AttributeError:
            # Fallback for non-Windows platforms
            pass
        dlg.destroy()
    ctk.CTkButton(btn_row, text="📂  Ouvrir le dossier", width=160,
                  command=_open).pack(side="left", padx=(0, 10))
    ctk.CTkButton(btn_row, text="Fermer", width=100,
                  fg_color="gray40", hover_color="gray30",
                  command=dlg.destroy).pack(side="left")


def _apply_zebra(path: str):
    """Apply zebra striping (alternating row colors + styled header) to an xlsx file."""
    import os
    import logging
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import PatternFill, Font
        wb = load_workbook(path)
        fill_hdr  = PatternFill(fill_type="solid", fgColor="4472C4")
        fill_even = PatternFill(fill_type="solid", fgColor="DCE6F1")
        fill_odd  = PatternFill(fill_type="solid", fgColor="FFFFFF")
        font_hdr  = Font(bold=True, color="FFFFFF")
        for ws in wb.worksheets:
            ncols = ws.max_column or 1
            nrows = ws.max_row or 1
            for c in range(1, ncols + 1):
                cell = ws.cell(1, c)
                cell.fill = fill_hdr
                cell.font = font_hdr
            for r in range(2, nrows + 1):
                rf = fill_even if (r % 2 == 0) else fill_odd
                for c in range(1, ncols + 1):
                    ws.cell(r, c).fill = rf
        tmp_path = path + ".tmp"
        wb.save(tmp_path)
        os.replace(tmp_path, path)
    except Exception as e:
        logging.exception(f"Erreur lors de l'application du style zébré sur {path}")


# ── FilePanel ─────────────────────────────────────────────────────────────────

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

    # ── Public ───────────────────────────────────────────────────────────────

    def get_column(self) -> str:
        return self.col_var.get()

    # ── Internal ─────────────────────────────────────────────────────────────

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
                    import fastexcel
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
        """Re-parse the selected sheet and update preview / columns."""
        if self._xls is None:
            return
        try:
            df = pl.read_excel(self._xls, sheet_name=sheet_name)
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


# ── Main Application ──────────────────────────────────────────────────────────

class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Outil de Filtrage Excel (Polars)")
        self.geometry("1150x860")
        self.minsize(900, 700)
        
        # Thème Moderne (Flat Dark)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.configure(fg_color="#181825") # Deep background
        
        self._result_df: pl.DataFrame | None = None
        self._notfound_df: pl.DataFrame | None = None
        self._en_found_df: pl.DataFrame | None = None
        self._col_vars:  dict[str, BooleanVar] = {}
        self._pdf_files: dict[str, str] = {} # {filename: path}
        self._op_running = False  # guard against concurrent runs
        self._cancel_requested = False
        self._build_ui()
        self.after(10, lambda: self.state("zoomed"))

    # =========================================================================
    # UI construction
    # =========================================================================

    def _build_ui(self):
        self._build_toolbar()
        self._build_navbar()
        scroll = ctk.CTkScrollableFrame(self, fg_color="#181825")
        scroll.pack(fill="both", expand=True, padx=10, pady=6)
        self._scroll = scroll
        self._build_files_section()
        self._build_operation_section()
        # Run button
        self._run_btn = ctk.CTkButton(
            scroll, text="▶  Lancer l'opération", height=44,
            fg_color="#a6e3a1", text_color="#181825", hover_color="#94cc90",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._run,
        )
        self._run_btn.pack(fill="x", padx=4, pady=(8, 4))
        
        # Stop button (only active when running)
        self._stop_btn = ctk.CTkButton(
            scroll, text="⏹  Arrêter l'opération", height=44,
            fg_color="#f38ba8", text_color="#181825", hover_color="#eba0ac",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._cancel_operation,
            state="disabled"
        )
        self._stop_btn.pack(fill="x", padx=4, pady=(0, 8))

        self._build_result_section()

    # ── Toolbar ───────────────────────────────────────────────────────────────

    def _build_navbar(self):
        """Fixed operation-selector bar pinned below the toolbar."""
        nav = ctk.CTkFrame(self, corner_radius=0, fg_color="#1e1e2e")
        nav.pack(fill="x", side="top")

        self._op_desc_var = StringVar(value=OP_DESC[OPERATIONS[0]])
        ctk.CTkLabel(
            nav, textvariable=self._op_desc_var,
            text_color="#89b4fa",
            font=ctk.CTkFont(size=11, slant="italic"),
        ).pack(anchor="w", padx=14, pady=(6, 0))

        self._op_var = StringVar(value=OPERATIONS[0])

        # Split into two rows so all 9 operations are always visible
        ROW1 = OPERATIONS[:5]   # Soustraire, Fusionner, Intersecter, Enrichir, Combiner
        ROW2 = OPERATIONS[5:]   # Extraire, Split (Sheets), Split (Columns), Excel to PDF

        def _on_row1(val):
            self._op_var.set(val)
            self._seg2.set("")   # deselect row 2
            self._op_changed(val)

        def _on_row2(val):
            self._op_var.set(val)
            self._seg1.set("")   # deselect row 1
            self._op_changed(val)

        self._seg1 = ctk.CTkSegmentedButton(
            nav, values=ROW1,
            selected_color="#1e66f5",
            selected_hover_color="#1142a1",
            font=ctk.CTkFont(size=13, weight="bold"),
            command=_on_row1,
        )
        self._seg1.set(OPERATIONS[0])
        self._seg1.pack(fill="x", padx=10, pady=(4, 2))

        self._seg2 = ctk.CTkSegmentedButton(
            nav, values=ROW2,
            selected_color="#1e66f5",
            selected_hover_color="#1142a1",
            font=ctk.CTkFont(size=13, weight="bold"),
            command=_on_row2,
        )
        self._seg2.set("")   # nothing selected in row 2 initially
        self._seg2.pack(fill="x", padx=10, pady=(0, 8))

    def _build_toolbar(self):
        tb = ctk.CTkFrame(self, height=50, corner_radius=0, fg_color="#11111b")
        tb.pack(fill="x", side="top")
        tb.pack_propagate(False)

        ctk.CTkLabel(
            tb, text="⚙  Outil de Filtrage Excel",
            text_color="#cdd6f4",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(side="left", padx=16)

        # right side (pack right-to-left)
        ctk.CTkButton(tb, text="Tout effacer", width=105,
                      fg_color="#f38ba8", text_color="#181825", hover_color="#d97d97",
                      command=self._clear_all).pack(side="right", padx=12)

    # ── Section 1: Files ──────────────────────────────────────────────────────

    def _build_files_section(self):
        sec = ctk.CTkFrame(self._scroll, fg_color="#313244", corner_radius=10, border_width=1, border_color="#45475a")
        sec.pack(fill="x", padx=4, pady=6)
        self._files_sec = sec          # keep reference for hide/show
        ctk.CTkLabel(sec, text="1 · Charger les fichiers",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     ).pack(anchor="w", padx=10, pady=(8, 4))

        self._pa = FilePanel(sec, "Fichier A  (principal)",   on_load=self._file_loaded)
        self._pb = FilePanel(sec, "Fichier B  (secondaire)", on_load=self._file_loaded)
        self._pa.pack(fill="x", padx=6, pady=(0, 4))
        self._pb.pack(fill="x", padx=6, pady=(0, 8))

    # ── Section 2: Operation ──────────────────────────────────────────────────

    def _build_operation_section(self):
        self._op_sec = ctk.CTkFrame(self._scroll, fg_color="#313244", corner_radius=10, border_width=1, border_color="#45475a")
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
            "pdf":      self._build_pdf_panel(host),
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
            text_color="#a6adc8", font=ctk.CTkFont(size=11),
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
                      fg_color="#89b4fa", text_color="#181825", hover_color="#74a1e9",
                      command=self._mi_multi_browse).pack(side="left", padx=(12, 4))
        ctk.CTkButton(mi_btn_row, text="Retirer", width=80,
                      fg_color="#f38ba8", text_color="#181825", hover_color="#d97d97",
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
        self._mi_multi_files: dict[str, pl.DataFrame] = {}

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

        # Hidden by default – shown only for Intersecter
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

        self._ex_filters_frame = ctk.CTkFrame(f, fg_color="transparent")
        self._ex_filters_frame.pack(fill="x", padx=4, pady=(4, 4))
        self._ex_filters = []

        btn_row = ctk.CTkFrame(f, fg_color="transparent")
        btn_row.pack(fill="x", padx=8, pady=(4, 4))
        ctk.CTkButton(btn_row, text="➕ Ajouter un filtre", width=140,
                      command=self._ex_add_filter).pack(side="left")

        self._ex_add_filter()

        return f

    def _ex_add_filter(self):
        row_frame = ctk.CTkFrame(self._ex_filters_frame)
        row_frame.pack(fill="x", pady=2)
        
        logic_var = StringVar(value="ET")
        logic_menu = ctk.CTkOptionMenu(row_frame, variable=logic_var, values=["ET", "OU"], width=60)
        
        cols = list(self._pa.df.columns) if getattr(self, "_pa", None) and self._pa.df is not None else ["— charger Fichier A —"]
        
        col_var = StringVar(value=cols[0] if cols else "")
        col_menu = ctk.CTkOptionMenu(row_frame, variable=col_var, values=cols, width=200)
        col_menu.pack(side="left", padx=4, pady=4)
        
        cond_var = StringVar(value=CONDITIONS[0])
        cond_menu = ctk.CTkOptionMenu(row_frame, variable=cond_var, values=CONDITIONS, width=150)
        cond_menu.pack(side="left", padx=4, pady=4)
        
        val_var = StringVar()
        val_entry = ctk.CTkEntry(row_frame, textvariable=val_var, width=160, placeholder_text="valeur …")
        val_entry.pack(side="left", padx=4, pady=4)
        
        case_var = BooleanVar(value=False)
        case_chk = ctk.CTkCheckBox(row_frame, text="Aa", variable=case_var, width=50)
        case_chk.pack(side="left", padx=8, pady=4)
        
        def _cond_changed(val):
            if val in ("est vide", "n'est pas vide"):
                val_entry.configure(state="disabled")
            else:
                val_entry.configure(state="normal")
        cond_menu.configure(command=_cond_changed)
        
        del_btn = ctk.CTkButton(row_frame, text="✕", width=30, fg_color="#f38ba8", text_color="#181825", hover_color="#d97d97")
        del_btn.pack(side="left", padx=4, pady=4)
        
        filter_dict = {
            "frame": row_frame,
            "logic_var": logic_var,
            "logic_menu": logic_menu,
            "col_var": col_var,
            "col_menu": col_menu,
            "cond_var": cond_var,
            "val_var": val_var,
            "case_var": case_var,
            "val_entry": val_entry
        }
        
        def _remove():
            row_frame.destroy()
            if filter_dict in self._ex_filters:
                self._ex_filters.remove(filter_dict)
            if not self._ex_filters:
                self._ex_add_filter()
            else:
                self._ex_refresh_filter_ui()
        del_btn.configure(command=_remove)
        
        self._ex_filters.append(filter_dict)
        self._ex_refresh_filter_ui()

    def _ex_refresh_filter_ui(self):
        for i, fil in enumerate(self._ex_filters):
            if i == 0:
                fil["logic_menu"].pack_forget()
            else:
                fil["logic_menu"].pack(side="left", padx=(4, 0), pady=4, before=fil["col_menu"])

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
            text_color="#a6adc8", font=ctk.CTkFont(size=11),
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
                     text_color="#a6adc8", font=ctk.CTkFont(size=10)).pack(side="left", padx=8)

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
                      fg_color="#89b4fa", text_color="#181825", hover_color="#74a1e9",
                      command=self._split_browse).pack(side="left")
        ctk.CTkButton(file_row, text="Tout retirer", width=130,
                      fg_color="#f38ba8", text_color="#181825", hover_color="#d97d97",
                      command=self._split_remove).pack(side="left", padx=8)

        # Internal storage: path → {sheet_name: row_count, …}
        self._sp_files: dict[str, dict[str, int]] = {}

        # ── Sheet selection checkboxes ────────────────────────────────────
        sheet_sel_host = ctk.CTkFrame(f)
        sheet_sel_host.pack(fill="x", pady=(4, 4))
        ctk.CTkLabel(sheet_sel_host, text="Onglets à exporter :",
                     font=ctk.CTkFont(size=12)).pack(anchor="w", padx=8, pady=(4, 2))
        self._sp_sheet_scroll = ctk.CTkScrollableFrame(
            sheet_sel_host, height=70, orientation="horizontal")
        self._sp_sheet_scroll.pack(fill="x", padx=8, pady=(0, 6))
        self._sp_sheet_vars: dict[str, BooleanVar] = {}   # "path||sheet" → var

        self._sp_name_var = StringVar(value="{fichier}_{onglet}")

        # ── Output dir ───────────────────────────────────────────────────
        out_row = ctk.CTkFrame(f, fg_color="transparent")
        out_row.pack(fill="x", pady=(4, 4))
        ctk.CTkLabel(out_row, text="Dossier de sortie :").pack(side="left", padx=(0, 6))
        self._sp_dir_var = StringVar(value=str(DEFAULT_OUT))
        ctk.CTkEntry(out_row, textvariable=self._sp_dir_var, width=260,
                     placeholder_text="(même dossier que le fichier source)").pack(side="left")
        ctk.CTkButton(out_row, text="…", width=36,
                      command=self._split_pick_dir).pack(side="left", padx=4)

        # ── Zebra mode ───────────────────────────────────────────────────
        sp_zebra_row = ctk.CTkFrame(f, fg_color="transparent")
        sp_zebra_row.pack(fill="x", pady=(2, 2))
        self._sp_zebra_var = BooleanVar(value=False)
        ctk.CTkCheckBox(
            sp_zebra_row,
            text="Mode zébré  (colorer les lignes en alternance)",
            variable=self._sp_zebra_var,
        ).pack(side="left")

        # ── Progress bar (hidden until a run starts) ─────────────────────
        self._sp_prog_frame = ctk.CTkFrame(f, fg_color="transparent")
        self._sp_prog_frame.pack(fill="x", padx=4, pady=(2, 4))
        self._sp_prog_label = ctk.CTkLabel(
            self._sp_prog_frame, text="",
            font=ctk.CTkFont(size=11))
        self._sp_prog_label.pack(anchor="w", padx=4)
        self._sp_prog_bar = ctk.CTkProgressBar(self._sp_prog_frame, height=14, progress_color="#a6e3a1")
        self._sp_prog_bar.set(0)
        self._sp_prog_bar.pack(fill="x", padx=4, pady=(2, 4))
        self._sp_prog_frame.pack_forget()

        return f

    # ── Splitcol panel ────────────────────────────────────────────────────────

    def _build_splitcol_panel(self, parent) -> ctk.CTkFrame:
        """Panel for Scinder (Colonnes): single-file column-based split."""
        f = ctk.CTkFrame(parent, fg_color="transparent")

        # ── File picker row ──────────────────────────────────────────────────
        file_row = ctk.CTkFrame(f, fg_color="transparent")
        file_row.pack(fill="x", pady=(4, 2))
        ctk.CTkButton(file_row, text="Parcourir …", width=130,
                      fg_color="#89b4fa", text_color="#181825", hover_color="#74a1e9",
                      command=self._sc_browse).pack(side="left")
        self._sc_fname_var = StringVar(value="Aucun fichier sélectionné")
        ctk.CTkLabel(file_row, textvariable=self._sc_fname_var,
                     text_color="#a6adc8", font=ctk.CTkFont(size=11)
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
        self._sc_df: pl.DataFrame | None = None

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
                      fg_color="#f38ba8", text_color="#181825", hover_color="#d97d97",
                      command=self._sc_remove_split_col).pack(side="left")

        # ── Output options row ───────────────────────────────────────────────
        opt_row = ctk.CTkFrame(f, fg_color="transparent")
        opt_row.pack(fill="x", pady=(2, 2))
        ctk.CTkLabel(opt_row, text="Dossier de sortie :").pack(side="left", padx=(0, 6))
        self._sc_dir_var = StringVar(value=str(DEFAULT_OUT))
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
        self._sc_prog_bar = ctk.CTkProgressBar(self._sc_prog_frame, height=14, progress_color="#a6e3a1")
        self._sc_prog_bar.set(0)
        self._sc_prog_bar.pack(fill="x", padx=4, pady=(2, 4))
        self._sc_prog_frame.pack_forget()

        return f

    # ── Splitcol helpers ──────────────────────────────────────────────────────

    def _sc_browse(self):
        path = filedialog.askopenfilename(
            title="Sélectionner le fichier à scinder",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tous", "*.*")],
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
                    xls_obj = fastexcel.read_excel(path)
                    sheets = xls_obj.sheet_names
                    df = operations.cleanup_floats(xls_obj.load_sheet(sheets[0]).to_polars())
                else:
                    df = load_file(path)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Erreur", str(e)))
                self.after(0, lambda: self._sc_fname_var.set("Aucun fichier sélectionné"))
                return

            def _apply():
                if old_xls is not None:
                    try:
                        if hasattr(old_xls, "close"):
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
                        if hasattr(xls_obj, "close"):
                            xls_obj.close()
                    except Exception:
                        pass
                self._sc_accept(path, df)
            self.after(0, _apply)

        threading.Thread(target=_load, daemon=True).start()

    def _sc_accept(self, path: str, df: pl.DataFrame):
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
            df = operations.cleanup_floats(self._sc_xls.load_sheet(sheet_name).to_polars())
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

    # ── PDF Reporting helpers ─────────────────────────────────────────────────

    def _build_pdf_panel(self, parent) -> ctk.CTkFrame:
        """Standalone panel to configure many-to-PDF reports."""
        f = ctk.CTkFrame(parent, fg_color="transparent")

        # ── Group 1: Source Folder ──────────────────────────────────────────
        src_fr = ctk.CTkFrame(f)
        src_fr.pack(fill="x", padx=4, pady=(4, 6))
        
        btn_row = ctk.CTkFrame(src_fr, fg_color="transparent")
        btn_row.pack(fill="x", padx=10, pady=(10, 6))
        
        ctk.CTkLabel(btn_row, text="Dossier source :", 
                     font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        
        ctk.CTkButton(btn_row, text="Tout décocher", width=100, height=26,
                      fg_color="#f38ba8", text_color="#181825", hover_color="#d97d97",
                      command=lambda: self._pdf_multi_toggle_all(False)).pack(side="right", padx=4)
        ctk.CTkButton(btn_row, text="Tout cocher", width=100, height=26,
                      fg_color="#a6e3a1", text_color="#181825", hover_color="#94cc90",
                      command=lambda: self._pdf_multi_toggle_all(True)).pack(side="right", padx=4)
        ctk.CTkButton(btn_row, text="Sélectionner Dossier …", width=170,
                      fg_color="#89b4fa", text_color="#181825", hover_color="#74a1e9",
                      command=self._pdf_multi_browse).pack(side="right", padx=10)

        self._pdf_tree_fr = ctk.CTkFrame(src_fr)
        self._pdf_tree_fr.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self._pdf_tree_fr.columnconfigure(0, weight=1)
        self._pdf_tree_fr.rowconfigure(0, weight=1)
        
        # Columns: 'sel' for checkbox, 'path' for filename
        self._pdf_tree = ttk.Treeview(self._pdf_tree_fr, columns=("sel", "path"), height=6, show="headings")
        self._pdf_tree.heading("sel", text="Sél.")
        self._pdf_tree.heading("path", text="Fichier Excel")
        self._pdf_tree.column("sel", width=40, anchor="center")
        self._pdf_tree.column("path", width=400, stretch=True)
        
        self._pdf_tree.bind("<ButtonRelease-1>", self._pdf_on_tree_click)
        
        pdf_vsb = ttk.Scrollbar(self._pdf_tree_fr, orient="vertical", command=self._pdf_tree.yview)
        self._pdf_tree.configure(yscrollcommand=pdf_vsb.set)
        self._pdf_tree.grid(row=0, column=0, sticky="nsew")
        pdf_vsb.grid(row=0, column=1, sticky="ns")

        # Selection state: set of filename keys
        self._pdf_selected = set()
        self._pdf_files = {} # name -> full_path

        # ── Group 2: Report Content ─────────────────────────────────────────
        meta_fr = ctk.CTkFrame(f)
        meta_fr.pack(fill="x", padx=4, pady=2)
        
        r1 = ctk.CTkFrame(meta_fr, fg_color="transparent")
        r1.pack(fill="x", padx=10, pady=(10, 4))
        ctk.CTkLabel(r1, text="Titre du rapport :", width=120, anchor="w").pack(side="left")
        self._pdf_title_var = StringVar(value="Rapport d'Activité")
        ctk.CTkEntry(r1, textvariable=self._pdf_title_var, width=300).pack(side="left", fill="x", expand=True)

        r2 = ctk.CTkFrame(meta_fr, fg_color="transparent")
        r2.pack(fill="x", padx=10, pady=(4, 10))
        ctk.CTkLabel(r2, text="En-tête secondaire :", width=120, anchor="w").pack(side="left")
        self._pdf_header_var = StringVar(value="Document généré automatiquement")
        ctk.CTkEntry(r2, textvariable=self._pdf_header_var, width=300).pack(side="left", fill="x", expand=True)

        # ── Group 3: Style & Settings ───────────────────────────────────────
        style_fr = ctk.CTkFrame(f)
        style_fr.pack(fill="x", padx=4, pady=6)

        r3 = ctk.CTkFrame(style_fr, fg_color="transparent")
        r3.pack(fill="x", padx=10, pady=(10, 4))
        
        ctk.CTkLabel(r3, text="Orientation :", width=100, anchor="w").pack(side="left")
        self._pdf_orient_var = StringVar(value="Portrait")
        ctk.CTkOptionMenu(r3, variable=self._pdf_orient_var, values=["Portrait", "Paysage"], width=140).pack(side="left")

        ctk.CTkLabel(r3, text="Langue / Sens :", width=110, anchor="e").pack(side="left", padx=(20, 10))
        self._pdf_dir_var = StringVar(value="Français (LTR)")
        ctk.CTkOptionMenu(r3, variable=self._pdf_dir_var, values=["Français (LTR)", "Arabe (RTL)"], width=160).pack(side="left")

        # ── Output ──────────────────────────────────────────────────────────
        # Removed "Dossier de sortie" as per user request to save in the source folder.
        
        # ── Progress bar ─────────────────────────────────────────────
        self._pdf_prog_frame = ctk.CTkFrame(f, fg_color="transparent")
        self._pdf_prog_frame.pack(fill="x", padx=4, pady=(2, 4))
        self._pdf_prog_label = ctk.CTkLabel(
            self._pdf_prog_frame, text="",
            font=ctk.CTkFont(size=11))
        self._pdf_prog_label.pack(anchor="w", padx=4)
        self._pdf_prog_bar = ctk.CTkProgressBar(self._pdf_prog_frame, height=14, progress_color="#a6e3a1")
        self._pdf_prog_bar.set(0)
        self._pdf_prog_bar.pack(fill="x", padx=4, pady=(2, 4))
        self._pdf_prog_frame.pack_forget()

        return f

    def _pdf_multi_browse(self):
        d = filedialog.askdirectory(title="Choisir le dossier contenant les Excels")
        if not d: return
        self._pdf_files.clear()
        self._pdf_selected.clear()
        p = pathlib.Path(d)
        for ext in ("*.xlsx", "*.xls"):
            for f in p.glob(ext):
                name = f.name
                self._pdf_files[name] = str(f)
                self._pdf_selected.add(name)
        self._pdf_multi_refresh_tree()

    def _pdf_multi_refresh_tree(self):
        self._pdf_tree.delete(*self._pdf_tree.get_children())
        for name in sorted(self._pdf_files.keys()):
            status = "☑" if name in self._pdf_selected else "☐"
            self._pdf_tree.insert("", "end", iid=name, values=(status, name))

    def _pdf_on_tree_click(self, event):
        item = self._pdf_tree.identify_row(event.y)
        if not item: return
        # Toggle
        if item in self._pdf_selected:
            self._pdf_selected.remove(item)
        else:
            self._pdf_selected.add(item)
        self._pdf_multi_refresh_tree()

    def _pdf_multi_toggle_all(self, state: bool):
        if state:
            self._pdf_selected = set(self._pdf_files.keys())
        else:
            self._pdf_selected.clear()
        self._pdf_multi_refresh_tree()

    # Removed _pdf_pick_dir as output is now in source folder

    def _op_pdf_report(self):
        if not self._pdf_selected:
            messagebox.showwarning("Aucun fichier", "Sélectionnez au moins un fichier dans la liste.")
            self._set_running(False)
            return

        title = self._pdf_title_var.get()
        subtitle = self._pdf_header_var.get()
        orient = self._pdf_orient_var.get()
        direc = self._pdf_dir_var.get()

        files_to_process = [self._pdf_files[name] for name in sorted(self._pdf_selected)]

        # F-10: Pre-run file existence check — warn about missing files before starting
        missing = [p for p in files_to_process if not pathlib.Path(p).exists()]
        if missing:
            missing_names = "\n".join(pathlib.Path(p).name for p in missing[:10])
            suffix = f"\n… et {len(missing) - 10} autre(s)" if len(missing) > 10 else ""
            proceed = messagebox.askyesno(
                "Fichiers introuvables",
                f"Les fichiers suivants sont introuvables sur le disque :\n\n"
                f"{missing_names}{suffix}\n\n"
                f"Continuer avec les fichiers restants ?"
            )
            if not proceed:
                self._set_running(False)
                return
            files_to_process = [p for p in files_to_process if pathlib.Path(p).exists()]
            if not files_to_process:
                messagebox.showwarning("Aucun fichier valide", "Aucun fichier sélectionné n'existe sur le disque.")
                self._set_running(False)
                return

        # Build list of (file_path, sheet_name)
        work_list = []
        for path in files_to_process:
            try:
                xls = fastexcel.read_excel(path)
                for s in xls.sheet_names:
                    work_list.append((path, s))
            except Exception:
                continue

        total_sheets = len(work_list)
        if total_sheets == 0:
            messagebox.showwarning("Aucun onglet", "Aucun onglet Excel n'a été trouvé dans les fichiers sélectionnés.")
            self._set_running(False)
            return

        # Show progress bar
        self._pdf_prog_frame.pack(fill="x", padx=4, pady=(2, 4))
        self._pdf_prog_bar.set(0)
        self._pdf_prog_label.configure(text=f"Préparation : 0 / {total_sheets} onglets ...")
        self.update_idletasks()

        def _worker():
            done_sheets = 0   # total processed (including skipped)
            created = 0        # actually generated PDFs (ISSUE-026)
            errors = []
            last_dest_dir = None  # track the folder of the last successfully created PDF

            for path, sname in work_list:
                if self._cancel_requested:
                    self.after(0, lambda: self._log("Génération PDF annulée par l'utilisateur."))
                    break

                try:
                    p = pathlib.Path(path)
                    # Load exact sheet
                    xls = fastexcel.read_excel(path)
                    df = operations.cleanup_floats(xls.load_sheet(sname).to_polars())

                    if df.is_empty():
                        done_sheets += 1
                        continue

                    # Generate name: {file}_{sheet}.pdf — saved next to the source Excel
                    safe_sname = "".join([c if c.isalnum() else "_" for c in sname])
                    out_name = f"{p.stem}_{safe_sname}.pdf"
                    dest = p.parent / out_name

                    # FIX ISSUE-002: capture loop vars by value using default args
                    def _row_progress(current, total,
                                      _done=done_sheets, _p=p, _sname=sname):
                        # F-11: Clamp to 0.99 — final 1.0 is set only in _final()
                        pct = min(
                            (_done / total_sheets) + (current / total / total_sheets),
                            0.99
                        )
                        self.after(0, lambda: (
                            self._pdf_prog_bar.set(pct),
                            self._pdf_prog_label.configure(
                                text=f"File: {_p.name} | Sheet: {_sname} | Rows: {current}/{total}")
                        ))

                    pdf_generator.generate_report(
                        df=df,
                        output_path=str(dest),
                        title=title,
                        subtitle=subtitle,
                        orientation=orient,
                        direction=direc,
                        progress_callback=_row_progress
                    )
                    created += 1
                    last_dest_dir = str(p.parent)
                except Exception as e:
                    errors.append(f"{pathlib.Path(path).name} [{sname}]: {e}")
                    logging.exception("Erreur génération PDF")

                done_sheets += 1
                self.after(0, lambda d=done_sheets: (
                    self._pdf_prog_bar.set(d / total_sheets),
                    self._pdf_prog_label.configure(text=f"Terminé : {d} / {total_sheets} onglets")
                ))

            def _final():
                self._pdf_prog_bar.set(1)
                if errors:
                    msg = "\n".join(errors[:5]) + ("\n..." if len(errors) > 5 else "")
                    messagebox.showwarning("Terminé avec erreurs", f"Certains rapports ont échoué :\n{msg}")

                # ISSUE-026: use 'created', not 'done_sheets', for accurate count
                status_txt = "annulée" if self._cancel_requested else "terminée"
                summary = f"Génération {status_txt}. {created} rapport(s) créé(s) dans le(s) dossier(s) source(s)."
                self._pdf_prog_label.configure(text=f"✓ {summary}")
                self._stats_var.set(summary)
                self._log(f"[Excel to PDF] {summary}")
                
                if not self._cancel_requested and last_dest_dir:
                    _show_done_dialog(self, "Export PDF terminé", summary, last_dest_dir)
                
                self._set_running(False)

            self.after(0, _final)

        threading.Thread(target=_worker, daemon=True).start()

    # ── Split helpers ─────────────────────────────────────────────────────────

    def _split_browse(self):
        paths = filedialog.askopenfilenames(
            title="Sélectionner un ou plusieurs fichiers Excel",
            filetypes=[("Excel", "*.xlsx *.xls")],
        )
        if not paths:
            return
        for p in paths:
            if p in self._sp_files:
                continue
            try:
                xls = fastexcel.read_excel(p)
                sheets: dict[str, int] = {}
                for s in xls.sheet_names:
                    df = operations.cleanup_floats(xls.load_sheet(s).to_polars())
                    sheets[s] = len(df)
                self._sp_files[p] = sheets
            except Exception as e:
                messagebox.showerror("Erreur", f"{pathlib.Path(p).name}\n{e}")
        self._split_refresh_ui()

    def _split_remove(self):
        self._sp_files.clear()
        self._split_refresh_ui()

    def _split_refresh_ui(self):
        # Rebuild sheet checkboxes
        for w in self._sp_sheet_scroll.winfo_children():
            w.destroy()
        self._sp_sheet_vars.clear()
        for path, sheets in self._sp_files.items():
            for sname, nrows in sheets.items():
                key = f"{path}||{sname}"
                var = BooleanVar(value=True)
                self._sp_sheet_vars[key] = var
                label = f"{pathlib.Path(path).stem} — {sname}  ({nrows:,})"
                ctk.CTkCheckBox(
                    self._sp_sheet_scroll, text=label,
                    variable=var,
                    font=ctk.CTkFont(size=11),
                ).pack(side="left", padx=(0, 12))

    def _split_pick_dir(self):
        d = filedialog.askdirectory(title="Dossier de sortie")
        if d:
            self._sp_dir_var.set(d)

    # ── Combine helpers ───────────────────────────────────────────────────

    def _combine_browse(self):
        paths = filedialog.askopenfilenames(
            title="Sélectionner des fichiers à combiner",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tous", "*.*")],
        )
        if not paths:
            return
        for p in paths:
            if p in self._cb_files:
                continue
            try:
                ext = pathlib.Path(p).suffix.lower()
                if ext not in (".xlsx", ".xls"):
                    raise ValueError("Seuls les fichiers Excel sont pris en charge.")
                xls = fastexcel.read_excel(p)
                sheets = xls.sheet_names
                sheet_data: dict[str, pl.DataFrame] = {}
                for s in sheets:
                    sheet_data[s] = operations.cleanup_floats(xls.load_sheet(s).to_polars())
                self._cb_files[p] = sheet_data
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
            total = sum(len(df) for df in sheets.values())
            fid = self._cb_tree.insert(
                "", "end", iid=path,
                text=pathlib.Path(path).name,
                values=(len(sheets), f"{total:,}"),
            )
            for sname, df in sheets.items():
                self._cb_tree.insert(fid, "end",
                                     text=f"  ↳ {sname}",
                                     values=("", f"{len(df):,}"))
            self._cb_tree.item(fid, open=True)

    def _combine_refresh_cols(self):
        """Rebuild column checkboxes from union of all loaded files."""
        # Compute the new union of columns
        new_cols: list[str] = []
        for sheets in self._cb_files.values():
            for df in sheets.values():
                for col in df.columns:
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

    def _combine_toggle_all(self, state: bool):
        for v in self._cb_col_vars.values():
            v.set(state)

    def _build_combine_panel(self, parent) -> ctk.CTkFrame:
        """Standalone panel to combine multiple Excel files."""
        f = ctk.CTkFrame(parent, fg_color="transparent")

        # ── File picker ───────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(f, fg_color="transparent")
        btn_row.pack(fill="x", pady=(4, 4))
        ctk.CTkButton(btn_row, text="Ajouter des fichiers …", width=170,
                      fg_color="#89b4fa", text_color="#181825", hover_color="#74a1e9",
                      command=self._combine_browse).pack(side="left")
        ctk.CTkButton(btn_row, text="Retirer sélection", width=130,
                      fg_color="#f38ba8", text_color="#181825", hover_color="#d97d97",
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
        self._cb_files: dict[str, dict[str, pl.DataFrame]] = {}

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
        sec = ctk.CTkFrame(self._scroll, fg_color="#313244", corner_radius=10, border_width=1, border_color="#45475a")
        sec.pack(fill="both", expand=True, padx=4, pady=6)
        self._result_sec = sec          # keep reference for hide/show
        ctk.CTkLabel(sec, text="3 · Résultat & Export",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     ).pack(anchor="w", padx=10, pady=(8, 4))

        # Stats
        self._stats_var = StringVar(value="Lancez une opération pour voir les résultats.")
        ctk.CTkLabel(sec, textvariable=self._stats_var, text_color="#a6adc8",
                     font=ctk.CTkFont(size=12)).pack(anchor="w", padx=10)

        # Result Treeview count label
        self._rtree_count_var = StringVar(value="")
        ctk.CTkLabel(sec, textvariable=self._rtree_count_var,
                     font=ctk.CTkFont(size=11, slant="italic"),
                     text_color="#6c7086").pack(anchor="w", padx=10, pady=(0, 2))

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
                     text_color="#f9e2af").pack(anchor="w", padx=10, pady=(6, 2))

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
        
        _ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self._out_var = StringVar(value=str(DEFAULT_OUT / f"{_ts}.xlsx"))
        ctk.CTkEntry(exp, textvariable=self._out_var,
                      width=320, placeholder_text="(choisissez le chemin de sauvegarde …)").pack(side="left")
        ctk.CTkButton(exp, text="…", width=36,
                       command=self._pick_output).pack(side="left", padx=4)
        ctk.CTkButton(exp, text="Enregistrer en Excel",
                       command=self._export).pack(side="left", padx=(14, 4))

        # Zebra mode for export
        exp_zebra = ctk.CTkFrame(sec, fg_color="transparent")
        exp_zebra.pack(fill="x", padx=8, pady=(0, 4))
        self._exp_zebra_var = BooleanVar(value=False)
        ctk.CTkCheckBox(
            exp_zebra,
            text="Mode zébré  (colorer les lignes en alternance)",
            variable=self._exp_zebra_var,
        ).pack(side="left")

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
                
            if hasattr(self, "_ex_filters"):
                for fil in self._ex_filters:
                    fil["col_menu"].configure(values=a_cols)
                    if fil["col_var"].get() not in a_cols:
                        fil["col_var"].set(a_cols[0])

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
        elif val == "Split  (Sheets)":
            key = "split"
        elif val == "Split  (Columns)":
            key = "splitcol"
        elif val == "Excel to PDF":
            key = "pdf"
        else:
            key = "mi"
        self._show_sub(key)
        # Show column selector only for intersect
        if "Intersecter" in val:
            self._mi_col_frame.grid()
        else:
            self._mi_col_frame.grid_remove()
        # Hide file panels and result section for standalone operations
        is_standalone = val in ("Split  (Sheets)", "Combiner  (Multi-fichiers)",
                                "Split  (Columns)", "Excel to PDF")
        is_scinder = val in ("Split  (Sheets)", "Split  (Columns)", "Excel to PDF")
        if is_standalone:
            self._files_sec.pack_forget()
        else:
            siblings = self._scroll.winfo_children()
            if self._files_sec not in [w for w in siblings if w.winfo_ismapped()]:
                self._files_sec.pack(fill="x", padx=4, pady=6,
                                      before=self._op_sec)
        # Hide result section for Scinder operations (they output files directly)
        if is_scinder:
            self._result_sec.pack_forget()
        else:
            siblings = self._scroll.winfo_children()
            if self._result_sec not in [w for w in siblings if w.winfo_ismapped()]:
                self._result_sec.pack(fill="both", expand=True, padx=4, pady=6,
                                       after=self._run_btn)

    # (old _cond_changed block removed)

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
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tous", "*.*")],
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
        if hasattr(self, "_ex_filters"):
            for fil in list(self._ex_filters):
                fil["frame"].destroy()
            self._ex_filters.clear()
            self._ex_add_filter()
        self._mi_multi_files.clear()
        self._mi_multi_refresh_tree()
        self._mi_multi_refresh_key()
        self._mi_mode_var.set("Standard (1 A × 1 B)")
        self._mi_multi_frame.grid_remove()
        self._sp_files.clear()
        self._split_refresh_ui()
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
        # Reset PDF Reporting
        self._pdf_files.clear()
        self._pdf_multi_refresh_tree()
        self._pdf_title_var.set("Rapport d'Activité")
        self._pdf_header_var.set("Document généré automatiquement")
        self._pdf_orient_var.set("Portrait")
        self._pdf_dir_var.set("Français (LTR)")
        # F-08: _pdf_out_var was removed when the output-folder row was deleted.
        # Do NOT reference it here.
        
        self._stats_var.set("Lancez une opération pour voir les résultats.")
        self._out_var.set("")
        self._log("État réinitialisé.")

    # =========================================================================
    # Operations
    # =========================================================================

    def _set_running(self, running: bool):
        """Enable/disable the Run button to prevent concurrent operations.

        F-09: _cancel_requested is reset to False here, on the main thread,
        BEFORE any worker thread is spawned in _run()/_op_split()/_op_splitcol()
        /_op_pdf_report(). This ordering guarantee must be preserved — do not
        move _cancel_requested = False into worker threads.
        """
        # Assert we are on the main thread (Tkinter is single-threaded)
        assert threading.current_thread() is threading.main_thread(), (
            "_set_running must be called from the main thread only."
        )
        self._op_running = running
        if running:
            self._cancel_requested = False  # Must happen before worker thread starts
            self._run_btn.configure(state="disabled")
            self._stop_btn.configure(state="normal")
        else:
            self._run_btn.configure(state="normal")
            self._stop_btn.configure(state="disabled")

    def _cancel_operation(self):
        if self._op_running:
            self._cancel_requested = True
            self._stop_btn.configure(state="disabled")
            self._log("Annulation demandée...")
            messagebox.showinfo(
                "Annulation demandée",
                "L'opération s'arrêtera immédiatement après avoir terminé l'élément en cours."
            )

    def _run(self):
        # Guard against concurrent operations
        if self._op_running:
            return
        op = self._op_var.get()
        self._notfound_df = None
        self._hide_notfound()
        self._stats_var.set("Opération en cours …")
        self._set_running(True)
        self.update_idletasks()

        # Standalone operations manage their own threads — call directly (no double-bounce)
        if op == "Split  (Columns)":
            self._op_splitcol()
            return
        if op == "Split  (Sheets)":
            self._op_split()
            return
        if op == "Excel to PDF":
            self._op_pdf_report()
            return

        # ── Snapshot all shared data on the main thread before dispatching ──
        try:
            df_a = self._pa.df.clone() if self._pa.df is not None else None
            df_b = self._pb.df.clone() if self._pb.df is not None else None
            # Capture copies of all mutable state needed by workers
            op_snapshot = op
            mi_mode = self._mi_mode_var.get()
            mi_a_col = self._mi_a_var.get()
            mi_b_col = self._mi_b_var.get()
            mi_multi_files = {p: df.clone() for p, df in self._mi_multi_files.items()}
            mi_multi_key = self._mi_multi_key_var.get()
            mi_nf = self._mi_nf_var.get()
            mi_a_col_vars = {c: v.get() for c, v in self._mi_a_col_vars.items()}
            mi_b_col_vars = {c: v.get() for c, v in self._mi_b_col_vars.items()}
            mi_suffix = self._mi_suffix_var.get()
            mi_principal = self._mi_principal_var.get()
            ext_dedup = self._ext_dedup.get()
            ext_key = self._ext_key_var.get()
            en_a_col = self._en_a_var.get()
            en_b_col = self._en_b_var.get()
            en_cols = {c: v.get() for c, v in self._en_col_vars.items()}
            en_mode = self._en_mode_var.get()
            en_suffix = self._en_suffix_var.get()
            en_skip_empty = self._en_skip_empty.get()
            en_fill = self._en_fill_var.get()
            en_fill_val = self._en_fill_entry.get() if en_fill == "Valeur personnalisée" else ""
            ex_filters = [
                {
                    "col": f["col_var"].get(),
                    "cond": f["cond_var"].get(),
                    "val": f["val_var"].get(),
                    "case": f["case_var"].get(),
                    "logic": f["logic_var"].get(),
                }
                for f in self._ex_filters
            ] if hasattr(self, "_ex_filters") else []
            cb_files = {p: {s: df.clone() for s, df in sheets.items()} for p, sheets in self._cb_files.items()}
            cb_dedup = self._cb_dedup.get()
            cb_dedup_key = self._cb_dedup_key_var.get()
            cb_cols = {c: v.get() for c, v in self._cb_col_vars.items()}
        except Exception as snap_err:
            self._set_running(False)
            messagebox.showerror("Erreur", f"Impossible de lire les données : {snap_err}")
            return

        def _compute():
            notfound_result = None
            en_found_result  = None
            try:
                if "Soustraire" in op_snapshot:
                    result, notfound_result = operations.op_minus_or_intersect(
                        minus=True,
                        df_a=df_a, df_b=df_b,
                        mi_mode=mi_mode, mi_a_col=mi_a_col, mi_b_col=mi_b_col,
                        mi_multi_files=mi_multi_files, mi_multi_key=mi_multi_key,
                        mi_nf=mi_nf, mi_a_col_vars=mi_a_col_vars, mi_b_col_vars=mi_b_col_vars,
                        mi_suffix=mi_suffix, mi_principal=mi_principal,
                    )
                elif "Fusionner" in op_snapshot:
                    result = operations.op_extend(df_a=df_a, df_b=df_b, ext_dedup=ext_dedup, ext_key=ext_key)
                elif "Intersecter" in op_snapshot:
                    result, notfound_result = operations.op_minus_or_intersect(
                        minus=False,
                        df_a=df_a, df_b=df_b,
                        mi_mode=mi_mode, mi_a_col=mi_a_col, mi_b_col=mi_b_col,
                        mi_multi_files=mi_multi_files, mi_multi_key=mi_multi_key,
                        mi_nf=mi_nf, mi_a_col_vars=mi_a_col_vars, mi_b_col_vars=mi_b_col_vars,
                        mi_suffix=mi_suffix, mi_principal=mi_principal,
                    )
                elif "Enrichir" in op_snapshot:
                    result, en_found_result = operations.op_enrich(
                        df_a=df_a, df_b=df_b,
                        en_a_col=en_a_col, en_b_col=en_b_col,
                        en_cols=en_cols, en_mode=en_mode,
                        en_suffix=en_suffix, en_skip_empty=en_skip_empty,
                        en_fill=en_fill, en_fill_val=en_fill_val,
                    )
                elif "Extraire" in op_snapshot:
                    result = operations.op_extract(df=df_a, ex_filters=ex_filters)
                elif "Combiner" in op_snapshot:
                    result = operations.op_combine(
                        cb_files=cb_files, cb_dedup=cb_dedup,
                        cb_dedup_key=cb_dedup_key, cb_cols=cb_cols,
                    )
                else:
                    raise ValueError("Opération inconnue")
            except Exception as e:
                logging.exception("Erreur d'opération")
                self.after(0, lambda: messagebox.showerror("Erreur d'opération", str(e)))
                self.after(0, lambda: self._stats_var.set(
                    "Lancez une opération pour voir les résultats."))
                self.after(0, lambda: self._set_running(False))
                return

            def _show_result():
                self._result_df = result
                self._notfound_df = notfound_result
                self._en_found_df = en_found_result
                a_n = len(df_a) if df_a is not None else None
                b_n = len(df_b) if df_b is not None else None
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
                    f"[{op_snapshot.strip()}]  A={fmt(a_n)}  B={fmt(b_n)}  → {len(result):,} lignes"
                    + (f"  | non trouvés={nf_n:,}" if nf_n else "")
                    + (f"  | trouvés B={ef_n:,}" if ef_n else "")
                )
                if "Combiner" in op_snapshot:
                    all_cols = list(result.columns)
                    self._cb_dedup_key_menu.configure(values=["— toutes les colonnes —"] + all_cols)

                self._set_running(False)
            self.after(0, _show_result)

        threading.Thread(target=_compute, daemon=True).start()



    def _op_split(self):
        """Split selected sheets into individual files — threaded with progress."""
        if not self._sp_files:
            messagebox.showwarning("Aucun fichier",
                                  "Ajoutez au moins un fichier Excel à scinder.")
            self._set_running(False)
            return

        # Build set of selected sheets: {(path, sheet_name), ...}
        selected_sheets: set = set()
        for key, var in self._sp_sheet_vars.items():
            if var.get():
                path_str, sname = key.split("||", 1)
                selected_sheets.add((path_str, sname))
        if not selected_sheets:
            messagebox.showwarning("Aucun onglet",
                                  "Cochez au moins un onglet à exporter.")
            self._set_running(False)
            return

        # Snapshot all state on main thread
        pattern = self._sp_name_var.get().strip() or "{fichier}_{onglet}"
        out_dir_str = self._sp_dir_var.get().strip()
        base_dir = pathlib.Path(out_dir_str) if out_dir_str else DEFAULT_OUT
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        dest_dir = base_dir / ts
        total_sheets = len(selected_sheets)
        sp_files_copy = dict(self._sp_files)
        zebra = self._sp_zebra_var.get()

        # Show progress bar
        self._sp_prog_frame.pack(fill="x", padx=4, pady=(2, 4))
        self._sp_prog_bar.set(0)
        self._sp_prog_label.configure(text=f"0 / {total_sheets} onglet(s) — démarrage …")
        self.update_idletasks()

        def _worker():
            total_written = 0
            error_msg = None
            try:
                dest_dir.mkdir(parents=True, exist_ok=True)
                for path, sheets in sp_files_copy.items():
                    if self._cancel_requested:
                        break
                    src = pathlib.Path(path)
                    try:
                        xls = fastexcel.read_excel(path)
                    except Exception:
                        continue
                    for sname in xls.sheet_names:
                        if self._cancel_requested:
                            break
                        if (path, sname) not in selected_sheets:
                            continue
                        df = operations.cleanup_floats(xls.load_sheet(sname).to_polars())
                        fname = pattern.format(fichier=src.stem, onglet=sname)
                        fname = fname.replace("/", "_").replace("\\", "_")
                        out_path = dest_dir / f"{fname}.xlsx"
                        df.write_excel(out_path)
                        if zebra:
                            _apply_zebra(str(out_path))
                        total_written += 1
                        done = total_written
                        self.after(0, lambda d=done: (
                            self._sp_prog_bar.set(d / total_sheets),
                            self._sp_prog_label.configure(
                                text=f"{d} / {total_sheets} onglet(s) traité(s)")
                        ))
            except Exception as e:
                error_msg = str(e)
                logging.exception("Erreur _op_split")

            def _done():
                cancelled = self._cancel_requested and error_msg is None
                self._sp_prog_bar.set(1 if not cancelled else min(total_written / total_sheets, 0.99))
                if error_msg:
                    self._sp_prog_label.configure(text=f"Erreur : {error_msg}")
                    messagebox.showerror("Erreur de scission", error_msg)
                else:
                    summary = (
                        f"{total_written} fichier(s) créé(s) à partir de "
                        f"{len(sp_files_copy)} source(s)."
                    )
                    if cancelled:
                        summary = f"Scission annulée. {summary}"
                    self._sp_prog_label.configure(text=f"✓ {summary}")
                    self._stats_var.set(summary)
                    self._log(f"[Scinder]  {summary}")
                    if not cancelled:
                        _show_done_dialog(self, "Scission terminée", summary, str(dest_dir))
                self._set_running(False)
            self.after(0, _done)

        threading.Thread(target=_worker, daemon=True).start()

    def _op_splitcol(self):
        """Split file(s) by column values — optimised, threaded, with progress bar."""
        from copy import copy as _copy

        # ── 1. Validate inputs (still on main thread) ──────────────────────
        if self._sc_df is None:
            messagebox.showwarning("Aucun fichier", "Chargez un fichier dans Scinder (Colonnes).")
            self._set_running(False)
            return
        if not self._sc_split_cols:
            messagebox.showwarning("Colonnes manquantes",
                                   "Ajoutez au moins une colonne de découpe.")
            self._set_running(False)
            return
        for col in self._sc_split_cols:
            if col not in self._sc_df.columns:
                messagebox.showerror("Colonne introuvable",
                                     f"La colonne '{col}' n'existe pas dans le fichier chargé.")
                self._set_running(False)
                return

        keep_cols = [c for c, v in self._sc_keep_vars.items() if v.get()]
        if not keep_cols:
            messagebox.showwarning("Aucune colonne",
                                   "Cochez au moins une colonne à conserver.")
            self._set_running(False)
            return

        # Snapshot all state on main thread
        out_dir_str = self._sc_dir_var.get().strip()
        src_path = getattr(self, "_sc_source_path", "") or ""
        base_dir = pathlib.Path(out_dir_str) if out_dir_str else DEFAULT_OUT
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        out_dir = base_dir / ts

        zebra = self._sc_zebra_var.get()
        split_cols = list(self._sc_split_cols)
        all_sheets_mode = (
            getattr(self, "_sc_all_sheets_var", None) is not None
            and self._sc_all_sheets_var.get()
            and self._sc_xls is not None
        )
        sheet_names = list(self._sc_xls.sheet_names) if all_sheets_mode else None
        df_snapshot = self._sc_df.clone()

        # ── 2. Pre-compute total number of output files ─────────────────
        def _count_groups(df: pl.DataFrame) -> int:
            cols_present = [c for c in split_cols if c in df.columns]
            if not cols_present:
                return 0
            return df.select(cols_present).n_unique()

        if all_sheets_mode:
            total_groups = 0
            for sname in sheet_names:
                try:
                    df_s = operations.cleanup_floats(self._sc_xls.load_sheet(sname).to_polars())
                    total_groups += _count_groups(df_s)
                except Exception:
                    pass
        else:
            total_groups = _count_groups(df_snapshot)

        if total_groups == 0:
            messagebox.showwarning("Aucun groupe",
                                   "Aucun fichier à créer avec cette configuration.")
            self._set_running(False)
            return

        # ── 3. Show progress bar ────────────────────────────────────────
        self._sc_prog_frame.pack(fill="x", padx=4, pady=(2, 4))
        self._sc_prog_bar.set(0)
        self._sc_prog_label.configure(
            text=f"0 / {total_groups} fichier(s) — démarrage …")
        self.update_idletasks()

        # ── 4. Helpers (closures — no self needed in thread) ────────────
        def _safe_fname(keys):
            if not isinstance(keys, tuple):
                keys = (keys,)
            parts = []
            for k in keys:
                s = str(k).strip()
                if s.endswith(".0"):
                    try:
                        if float(s).is_integer():
                            s = s[:-2]
                    except ValueError:
                        pass
                for ch in '/\\:*?"<>|':
                    s = s.replace(ch, "_")
                parts.append(s)
            return ("_".join(parts) if any(parts) else "(vide)")[:200]

        def _safe_sheetname(name: str) -> str:
            for ch in '/\\:*?"<>|':
                name = name.replace(ch, "_")
            return name[:60]

        # ── 5. Fast split: write each group directly with Polars ───────
        def _process_fast(df_sheet, dest_dir, counter):
            """Fast path: write groups using Polars write_excel (no formatting copy)."""
            keep_present = [c for c in keep_cols if c in df_sheet.columns]
            cols_present = [c for c in split_cols if c in df_sheet.columns]
            if not keep_present or not cols_present:
                return counter

            dest_dir.mkdir(parents=True, exist_ok=True)
            groups = df_sheet.partition_by(cols_present, maintain_order=True, as_dict=True)

            for keys, group_df in groups.items():
                if self._cancel_requested:
                    break
                if isinstance(keys, str):
                    keys = (keys,)
                fname = _safe_fname(keys)
                out_path = dest_dir / f"{fname}.xlsx"
                group_df.select(keep_present).write_excel(str(out_path))
                counter += 1
                done = counter
                self.after(0, lambda d=done: (
                    self._sc_prog_bar.set(d / total_groups),
                    self._sc_prog_label.configure(
                        text=f"{d} / {total_groups} fichier(s) créé(s)")
                ))
            return counter

        # ── 5b. Zebra split: write with Polars then apply zebra fills ──
        def _process_zebra(df_sheet, dest_dir, counter):
            """Zebra path: write data with Polars, then apply zebra styling with openpyxl."""
            keep_present = [c for c in keep_cols if c in df_sheet.columns]
            cols_present = [c for c in split_cols if c in df_sheet.columns]
            if not keep_present or not cols_present:
                return counter

            dest_dir.mkdir(parents=True, exist_ok=True)
            groups = df_sheet.partition_by(cols_present, maintain_order=True, as_dict=True)

            for keys, group_df in groups.items():
                if self._cancel_requested:
                    break
                if isinstance(keys, str):
                    keys = (keys,)
                fname = _safe_fname(keys)
                out_path = dest_dir / f"{fname}.xlsx"
                sub = group_df.select(keep_present)
                sub.write_excel(str(out_path))
                _apply_zebra(str(out_path))
                counter += 1
                done = counter
                self.after(0, lambda d=done: (
                    self._sc_prog_bar.set(d / total_groups),
                    self._sc_prog_label.configure(
                        text=f"{d} / {total_groups} fichier(s) créé(s)")
                ))
            return counter

        # Select the appropriate processing function
        _process = _process_zebra if zebra else _process_fast

        # ── 6. Background worker ────────────────────────────────────────
        def _worker():

            total_written = 0
            error_msg = None
            try:
                out_dir.mkdir(parents=True, exist_ok=True)

                if all_sheets_mode:
                    for sname in sheet_names:
                        if self._cancel_requested:
                            break
                        try:
                            df_s = operations.cleanup_floats(pl.read_excel(src_path, sheet_name=sname))
                            
                        except Exception:
                            continue
                        sdir = out_dir / _safe_sheetname(sname)
                        total_written = _process(df_s, sdir, total_written)
                        if self._cancel_requested:
                            break
                else:
                    df_clean = df_snapshot
                    total_written = _process(df_clean, out_dir, total_written)
            except Exception as e:
                error_msg = str(e)
                logging.exception("Erreur _op_splitcol")

            def _done():
                cancelled = self._cancel_requested and error_msg is None
                self._sc_prog_bar.set(1 if not cancelled else min(total_written / total_groups, 0.99))
                if error_msg:
                    self._sc_prog_label.configure(text=f"Erreur : {error_msg}")
                    messagebox.showerror("Erreur de scission (colonnes)", error_msg)
                else:
                    summary = f"{total_written} fichier(s) créé(s) dans {out_dir}"
                    if cancelled:
                        summary = f"Scission annulée. {summary}"
                    self._sc_prog_label.configure(text=f"✓ {summary}")
                    self._stats_var.set(summary)
                    self._log(f"[Scinder Colonnes]  {total_written} fichier(s) → {out_dir}")
                    if not cancelled:
                        _show_done_dialog(self, "Scission terminée", summary, str(out_dir))
                self._set_running(False)
            self.after(0, _done)

        threading.Thread(target=_worker, daemon=True).start()


    # =========================================================================
    # Result table helpers
    # =========================================================================
    # =========================================================================
    # Result table helpers
    # =========================================================================

    def _fill_result_tree(self, df: pl.DataFrame, max_rows: int = 100):
        self._rtree.delete(*self._rtree.get_children())
        shown = min(len(df), max_rows)
        self._rtree_count_var.set(f"Affichage de {shown:,} ligne(s) sur {len(df):,}".replace(",", " "))
        self._rtree["columns"] = list(df.columns)
        for c in df.columns:
            self._rtree.heading(c, text=c)
            self._rtree.column(c, width=110, minwidth=60, stretch=False)
        for row in df.head(max_rows).rows():
            self._rtree.insert("", "end", values=row)

    def _refresh_col_checks(self, df: pl.DataFrame):
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

    def _show_notfound(self, df: pl.DataFrame):
        self._nf_stats_var.set(
            f"⚠  Non trouvés dans Fichier A : {len(df):,} lignes du Fichier B"
        )
        self._nf_tree.delete(*self._nf_tree.get_children())
        self._nf_tree["columns"] = list(df.columns)
        for c in df.columns:
            self._nf_tree.heading(c, text=c)
            self._nf_tree.column(c, width=110, minwidth=60, stretch=False)
        for row in df.head(100).rows():
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
            filetypes=[("Excel", "*.xlsx"), ("Tous", "*.*")],
        )
        if path:
            self._out_var.set(ensure_excel_output_path(path))

    def _export(self):
        if self._result_df is None:
            messagebox.showwarning("Aucun résultat", "Lancez d'abord une opération.")
            return
        path = ensure_excel_output_path(self._out_var.get())
        if not path:
            path = filedialog.asksaveasfilename(
                title="Enregistrer le résultat",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx"), ("Tous", "*.*")],
            )
            if not path:
                return
            path = ensure_excel_output_path(path)
        self._out_var.set(path)

        selected = [c for c, v in self._col_vars.items() if v.get()]
        if not selected:
            selected = list(self._result_df.columns)
        df_out = self._result_df.select(selected)

        has_extra = (
            (self._notfound_df is not None and not self._notfound_df.is_empty())
            or (self._en_found_df is not None and not self._en_found_df.is_empty())
        )

        try:
            pathlib.Path(path).parent.mkdir(parents=True, exist_ok=True)
            if has_extra:
                try:
                    import xlsxwriter
                except ImportError:
                    messagebox.showerror("Erreur", "Le module xlsxwriter est requis pour cette opération.")
                    return

                tmp_path = path + ".tmp"
                with xlsxwriter.Workbook(tmp_path) as workbook:
                    df_out.write_excel(workbook=workbook, worksheet="Résultat")
                    if self._notfound_df is not None and not self._notfound_df.is_empty():
                        self._notfound_df.write_excel(workbook=workbook, worksheet="Non trouvés")
                    if self._en_found_df is not None and not self._en_found_df.is_empty():
                        self._en_found_df.write_excel(workbook=workbook, worksheet="Trouvés dans B")
                os.replace(tmp_path, path)
            else:
                df_out.write_excel(path)
            self._log(
                f"Enregistré : {len(df_out):,} lignes × {len(selected)} colonnes  →  "
                f"{pathlib.Path(path).name}"
            )
            if self._exp_zebra_var.get():
                _apply_zebra(path)
            _show_done_dialog(self, "Enregistré", f"Fichier sauvegardé :\n{path}", path)
        except Exception as e:
            messagebox.showerror("Erreur d'enregistrement", str(e))

    # =========================================================================
    # History log
    # =========================================================================

    def _log(self, msg: str):
        logging.info(msg)
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self._log_box.configure(state="normal")
        self._log_box.insert("end", f"[{ts}]  {msg}\n")
        
        # ISSUE-019: Trim log to 200 lines
        lines = int(self._log_box.index("end-1c").split(".")[0])
        if lines > 200:
            self._log_box.delete("1.0", f"{lines - 200}.0")
            
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
        import sys
        sys.stderr.write(traceback.format_exc())
        from tkinter import messagebox as _mb
        import tkinter as _tk
        _r = _tk.Tk(); _r.withdraw()
        _mb.showerror("Erreur au démarrage", traceback.format_exc())
        _r.destroy()
