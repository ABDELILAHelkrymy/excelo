import threading
import pathlib
import datetime
import logging
from tkinter import messagebox
import polars as pl
import fastexcel
import operations
from export_utils import apply_zebra
import pdf_generator
from core.printer_service import get_default_printer, set_default_printer

def run_split_sheets(
    sp_files: dict,
    selected_sheets: set,
    pattern: str,
    out_dir_str: str,
    zebra: bool,
    progress_bar,
    progress_label,
    stats_var,
    log_func,
    show_done_func,
    set_running_func,
    cancel_check_func,
    after_func,
    default_out_path
):
    base_dir = pathlib.Path(out_dir_str) if out_dir_str else default_out_path
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    dest_dir = base_dir / ts
    total_sheets = len(selected_sheets)

    progress_bar.set(0)
    progress_label.configure(text=f"0 / {total_sheets} onglet(s) — démarrage …")

    def _worker():
        total_written = 0
        error_msg = None
        try:
            dest_dir.mkdir(parents=True, exist_ok=True)
            for path, sheets in sp_files.items():
                if cancel_check_func():
                    break
                src = pathlib.Path(path)
                try:
                    xls = fastexcel.read_excel(path)
                except Exception:
                    continue
                for sname in xls.sheet_names:
                    if cancel_check_func():
                        break
                    if (path, sname) not in selected_sheets:
                        continue
                    df = operations.cleanup_floats(xls.load_sheet(sname).to_polars())
                    fname = pattern.format(fichier=src.stem, onglet=sname)
                    fname = fname.replace("/", "_").replace("\\", "_")
                    out_path = dest_dir / f"{fname}.xlsx"
                    df.write_excel(out_path)
                    if zebra:
                        apply_zebra(str(out_path))
                    total_written += 1
                    done = total_written
                    after_func(0, lambda d=done: (
                        progress_bar.set(d / total_sheets),
                        progress_label.configure(text=f"{d} / {total_sheets} onglet(s) traité(s)")
                    ))
        except Exception as e:
            error_msg = str(e)
            logging.exception("Erreur _op_split")

        def _done():
            cancelled = cancel_check_func() and error_msg is None
            progress_bar.set(1 if not cancelled else min(total_written / total_sheets, 0.99))
            if error_msg:
                progress_label.configure(text=f"Erreur : {error_msg}")
                messagebox.showerror("Erreur de scission", error_msg)
            else:
                summary = f"{total_written} fichier(s) créé(s) à partir de {len(sp_files)} source(s)."
                if cancelled:
                    summary = f"Scission annulée. {summary}"
                progress_label.configure(text=f"✓ {summary}")
                stats_var.set(summary)
                log_func(f"[Scinder]  {summary}")
                if not cancelled:
                    show_done_func("Scission terminée", summary, str(dest_dir))
            set_running_func(False)
        after_func(0, _done)

    threading.Thread(target=_worker, daemon=True).start()


def run_split_columns(
    df_snapshot: pl.DataFrame,
    src_path: str,
    out_dir_str: str,
    split_cols: list[str],
    keep_cols: list[str],
    out_format: str,
    zebra: bool,
    pdf_dir: str,
    auto_print: bool,
    target_printer: str,
    all_sheets_mode: bool,
    sheet_names: list[str],
    total_groups: int,
    progress_bar,
    progress_label,
    stats_var,
    log_func,
    show_done_func,
    set_running_func,
    cancel_check_func,
    after_func,
    default_out_path
):
    base_dir = pathlib.Path(out_dir_str) if out_dir_str else default_out_path
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out_dir = base_dir / ts

    progress_bar.set(0)
    progress_label.configure(text=f"0 / {total_groups} fichier(s) — démarrage …")

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

    def _process_fast(df_sheet, dest_dir, counter):
        keep_present = [c for c in keep_cols if c in df_sheet.columns]
        cols_present = [c for c in split_cols if c in df_sheet.columns]
        if not keep_present or not cols_present:
            return counter
        dest_dir.mkdir(parents=True, exist_ok=True)
        groups = df_sheet.partition_by(cols_present, maintain_order=True, as_dict=True)
        for keys, group_df in groups.items():
            if cancel_check_func():
                break
            if isinstance(keys, str):
                keys = (keys,)
            fname = _safe_fname(keys)
            out_path = dest_dir / f"{fname}.xlsx"
            group_df.select(keep_present).write_excel(str(out_path))
            counter += 1
            done = counter
            after_func(0, lambda d=done: (
                progress_bar.set(d / total_groups),
                progress_label.configure(text=f"{d} / {total_groups} fichier(s) créé(s)")
            ))
        return counter

    def _process_zebra(df_sheet, dest_dir, counter):
        keep_present = [c for c in keep_cols if c in df_sheet.columns]
        cols_present = [c for c in split_cols if c in df_sheet.columns]
        if not keep_present or not cols_present:
            return counter
        dest_dir.mkdir(parents=True, exist_ok=True)
        groups = df_sheet.partition_by(cols_present, maintain_order=True, as_dict=True)
        for keys, group_df in groups.items():
            if cancel_check_func():
                break
            if isinstance(keys, str):
                keys = (keys,)
            fname = _safe_fname(keys)
            out_path = dest_dir / f"{fname}.xlsx"
            sub = group_df.select(keep_present)
            sub.write_excel(str(out_path))
            apply_zebra(str(out_path))
            counter += 1
            done = counter
            after_func(0, lambda d=done: (
                progress_bar.set(d / total_groups),
                progress_label.configure(text=f"{d} / {total_groups} fichier(s) créé(s)")
            ))
        return counter

    def _process_pdf(df_sheet, dest_dir, counter):
        keep_present = [c for c in keep_cols if c in df_sheet.columns]
        cols_present = [c for c in split_cols if c in df_sheet.columns]
        if not keep_present or not cols_present:
            return counter
        dest_dir.mkdir(parents=True, exist_ok=True)
        groups = df_sheet.partition_by(cols_present, maintain_order=True, as_dict=True)
        
        orig_printer = None
        if auto_print and target_printer != "Par défaut":
            orig_printer = get_default_printer()
            if target_printer != orig_printer:
                set_default_printer(target_printer)

        for keys, group_df in groups.items():
            if cancel_check_func():
                break
            if isinstance(keys, str):
                keys = (keys,)
            fname = _safe_fname(keys)
            out_path = dest_dir / f"{fname}.pdf"
            
            title = " - ".join(str(k) for k in keys) if keys else "Rapport"
            sub = group_df.select(keep_present)
            
            try:
                pdf_generator.generate_report(
                    df=sub,
                    output_path=str(out_path),
                    title=title,
                    subtitle="Extrait généré par Split (Colonnes)",
                    orientation="Paysage",
                    direction=pdf_dir
                )
                if auto_print:
                    import os
                    try:
                        os.startfile(str(out_path), "print")
                        import time
                        time.sleep(1.5)
                    except Exception as e:
                        logging.error(f"Échec de l'impression de {out_path}: {e}")
            except Exception as e:
                logging.exception(f"Erreur PDF pour {fname}")
                
            counter += 1
            done = counter
            after_func(0, lambda d=done: (
                progress_bar.set(d / total_groups),
                progress_label.configure(text=f"{d} / {total_groups} fichier(s) créé(s)")
            ))
            
        if auto_print and orig_printer and target_printer != "Par défaut":
            set_default_printer(orig_printer)
            
        return counter

    if out_format == "PDF (.pdf)":
        _process = _process_pdf
    else:
        _process = _process_zebra if zebra else _process_fast

    def _worker():
        total_written = 0
        error_msg = None
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
            if all_sheets_mode:
                for sname in sheet_names:
                    if cancel_check_func():
                        break
                    try:
                        df_s = operations.cleanup_floats(pl.read_excel(src_path, sheet_name=sname))
                    except Exception:
                        continue
                    sdir = out_dir / _safe_sheetname(sname)
                    total_written = _process(df_s, sdir, total_written)
                    if cancel_check_func():
                        break
            else:
                df_clean = df_snapshot
                total_written = _process(df_clean, out_dir, total_written)
        except Exception as e:
            error_msg = str(e)
            logging.exception("Erreur _op_splitcol")

        def _done():
            cancelled = cancel_check_func() and error_msg is None
            progress_bar.set(1 if not cancelled else min(total_written / total_groups, 0.99))
            if error_msg:
                progress_label.configure(text=f"Erreur : {error_msg}")
                messagebox.showerror("Erreur de scission (colonnes)", error_msg)
            else:
                summary = f"{total_written} fichier(s) créé(s) dans {out_dir}"
                if cancelled:
                    summary = f"Scission annulée. {summary}"
                progress_label.configure(text=f"✓ {summary}")
                stats_var.set(summary)
                log_func(f"[Scinder Colonnes]  {total_written} fichier(s) → {out_dir}")
                if not cancelled:
                    show_done_func("Scission terminée", summary, str(out_dir))
            set_running_func(False)
        after_func(0, _done)

    threading.Thread(target=_worker, daemon=True).start()

def run_split_lines(
    df: pl.DataFrame,
    source_path: str,
    out_dir_str: str,
    lines_per_file: int,
    zebra: bool,
    progress_bar,
    progress_label,
    stats_var,
    log_func,
    show_done_func,
    set_running_func,
    cancel_check_func,
    after_func,
    default_out_path
):
    import math
    base_dir = pathlib.Path(out_dir_str) if out_dir_str else default_out_path
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out_dir = base_dir / ts

    total_rows = len(df)
    total_files = math.ceil(total_rows / lines_per_file)

    progress_bar.set(0)
    progress_label.configure(text=f"0 / {total_files} fichier(s) — démarrage …")

    def _worker():
        written = 0
        error_msg = None
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
            src = pathlib.Path(source_path)
            base_name = src.stem

            for i in range(total_files):
                if cancel_check_func():
                    break
                chunk = df.slice(i * lines_per_file, lines_per_file)
                safe_name = f"{base_name}_part{i+1}"
                out_path = out_dir / f"{safe_name}.xlsx"
                
                chunk.write_excel(str(out_path))
                if zebra:
                    apply_zebra(str(out_path))
                    
                written += 1
                done = written
                after_func(0, lambda d=done: (
                    progress_bar.set(d / total_files),
                    progress_label.configure(text=f"{d} / {total_files} fichier(s) créé(s)")
                ))
        except Exception as e:
            error_msg = str(e)
            logging.exception("Erreur _op_splitline")

        def _done():
            cancelled = cancel_check_func() and error_msg is None
            progress_bar.set(1 if not cancelled else min(written / total_files, 0.99))
            if error_msg:
                progress_label.configure(text=f"Erreur : {error_msg}")
                messagebox.showerror("Erreur de scission (lignes)", error_msg)
            else:
                summary = f"{written} fichier(s) créé(s) dans {out_dir}"
                if cancelled:
                    summary = f"Scission annulée. {summary}"
                progress_label.configure(text=f"✓ {summary}")
                stats_var.set(summary)
                log_func(f"[Scinder Lignes]  {written} fichier(s) → {out_dir}")
                if not cancelled:
                    show_done_func("Scission par lignes terminée", summary, str(out_dir))
            set_running_func(False)
        after_func(0, _done)

    threading.Thread(target=_worker, daemon=True).start()

