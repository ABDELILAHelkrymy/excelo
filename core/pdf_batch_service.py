import threading
import pathlib
from tkinter import messagebox
import polars as pl
import fastexcel
import logging
import pdf_generator
import operations

def run_pdf_report_batch(
    files_to_process: list[str],
    title: str,
    subtitle: str,
    orient: str,
    direc: str,
    progress_bar,
    progress_label,
    stats_var,
    log_func,
    show_done_func,
    set_running_func,
    cancel_check_func,
    after_func
):
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
        set_running_func(False)
        return

    # Initialize UI elements from caller thread (usually main thread anyway)
    progress_bar.set(0)
    progress_label.configure(text=f"Préparation : 0 / {total_sheets} onglets ...")

    def _worker():
        done_sheets = 0   
        created = 0        
        errors = []
        last_dest_dir = None  

        for path, sname in work_list:
            if cancel_check_func():
                after_func(0, lambda: log_func("Génération PDF annulée par l'utilisateur."))
                break

            try:
                p = pathlib.Path(path)
                xls = fastexcel.read_excel(path)
                df = operations.cleanup_floats(xls.load_sheet(sname).to_polars())

                if df.is_empty():
                    done_sheets += 1
                    continue

                safe_sname = "".join([c if c.isalnum() else "_" for c in sname])
                out_name = f"{p.stem}_{safe_sname}.pdf"
                dest = p.parent / out_name

                def _row_progress(current, total, _done=done_sheets, _p=p, _sname=sname):
                    pct = min((_done / total_sheets) + (current / total / total_sheets), 0.99)
                    after_func(0, lambda: (
                        progress_bar.set(pct),
                        progress_label.configure(text=f"File: {_p.name} | Sheet: {_sname} | Rows: {current}/{total}")
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
            after_func(0, lambda d=done_sheets: (
                progress_bar.set(d / total_sheets),
                progress_label.configure(text=f"Terminé : {d} / {total_sheets} onglets")
            ))

        def _final():
            progress_bar.set(1)
            if errors:
                msg = "\n".join(errors[:5]) + ("\n..." if len(errors) > 5 else "")
                messagebox.showwarning("Terminé avec erreurs", f"Certains rapports ont échoué :\n{msg}")

            status_txt = "annulée" if cancel_check_func() else "terminée"
            summary = f"Génération {status_txt}. {created} rapport(s) créé(s) dans le(s) dossier(s) source(s)."
            progress_label.configure(text=f"✓ {summary}")
            stats_var.set(summary)
            log_func(f"[Excel to PDF] {summary}")
            
            if not cancel_check_func() and last_dest_dir:
                show_done_func("Export PDF terminé", summary, last_dest_dir)
            
            set_running_func(False)

        after_func(0, _final)

    threading.Thread(target=_worker, daemon=True).start()
