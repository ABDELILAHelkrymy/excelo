import pathlib
import os
import polars as pl
import customtkinter as ctk

def load_file(path: str) -> pl.DataFrame:
    import operations
    p = pathlib.Path(path)
    if p.suffix.lower() in (".xlsx", ".xls"):
        return operations.cleanup_floats(pl.read_excel(path))
    raise ValueError(f"Unsupported file type: {p.suffix}")

def fmt(n) -> str:
    """Format a row-count or return '—' for None."""
    return f"{n:,}" if isinstance(n, int) else "—"

def show_done_dialog(parent, title: str, message: str, open_path: str):
    """Show a completion dialog with 'Ouvrir' and 'Fermer' buttons."""
    w, h = 460, 150
    dlg = ctk.CTkToplevel(parent)
    dlg.title(title)
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
            pass
        dlg.destroy()
    ctk.CTkButton(btn_row, text="📂  Ouvrir le dossier", width=160,
                  command=_open).pack(side="left", padx=(0, 10))
    ctk.CTkButton(btn_row, text="Fermer", width=100,
                  fg_color="gray40", hover_color="gray30",
                  command=dlg.destroy).pack(side="left")
