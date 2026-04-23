import logging
import threading
import pathlib
import datetime
from tkinter import messagebox

def run_pdf_to_excel_batch(
    files_to_process,
    out_dir_str,
    merge_tables,
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
    try:
        import fitz
    except ImportError:
        messagebox.showerror("Dépendance manquante", "Veuillez installer PyMuPDF (fitz) avec la commande :\npip install PyMuPDF")
        set_running_func(False)
        return

    import polars as pl
    import re

    ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
    def clean_text(val):
        if isinstance(val, str):
            return ILLEGAL_CHARACTERS_RE.sub("", val)
        return val

    base_dir = pathlib.Path(out_dir_str) if out_dir_str else default_out_path
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out_dir = base_dir / f"PDF_to_Excel_{ts}"
    
    total_files = len(files_to_process)
    progress_bar.set(0)
    progress_label.configure(text=f"Préparation : 0 / {total_files} fichiers ...")

    def _worker():
        done_files = 0
        errors = []
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
            
            for pdf_path in files_to_process:
                if cancel_check_func():
                    break
                    
                path = pathlib.Path(pdf_path)
                try:
                    doc = fitz.open(str(path))
                    all_tables_data = []
                    
                    for page_num in range(len(doc)):
                        if cancel_check_func():
                            break
                        page = doc[page_num]
                        tabs = page.find_tables()
                        if not tabs.tables:
                            continue
                            
                        for i, tab in enumerate(tabs):
                            table_data = tab.extract()
                            if not table_data or len(table_data) < 2:
                                continue
                                
                            # Clean up the text data
                            cleaned_data = []
                            for row in table_data:
                                cleaned_row = [clean_text(cell) for cell in row]
                                cleaned_data.append(cleaned_row)
                                
                            if merge_tables:
                                if not all_tables_data:
                                    all_tables_data.extend(cleaned_data)
                                else:
                                    # Skip header if merging
                                    all_tables_data.extend(cleaned_data[1:])
                            else:
                                all_tables_data.append(cleaned_data)
                                
                    doc.close()
                    
                    if all_tables_data:
                        out_file = out_dir / f"{path.stem}.xlsx"
                        
                        if merge_tables:
                            df = pl.DataFrame(all_tables_data[1:], schema=all_tables_data[0], orient="row")
                            df.write_excel(str(out_file))
                        else:
                            # Save each table as a separate sheet
                            # Polars write_excel supports multiple dataframes as dict
                            sheets_dict = {}
                            for i, t_data in enumerate(all_tables_data):
                                if len(t_data) > 1:
                                    # Ensure unique column names if there are duplicates
                                    cols = []
                                    seen = set()
                                    for c in t_data[0]:
                                        c_name = str(c) if c else f"Col_{len(cols)}"
                                        if c_name in seen:
                                            c_name = f"{c_name}_{len(cols)}"
                                        seen.add(c_name)
                                        cols.append(c_name)
                                        
                                    df = pl.DataFrame(t_data[1:], schema=cols, orient="row")
                                    sheets_dict[f"Table_{i+1}"] = df
                            
                            if sheets_dict:
                                pl.DataFrame().write_excel(workbook=str(out_file)) # dummy to create
                                import xlsxwriter
                                with xlsxwriter.Workbook(str(out_file)) as workbook:
                                    for s_name, s_df in sheets_dict.items():
                                        s_df.write_excel(workbook=workbook, worksheet=s_name)

                except Exception as e:
                    errors.append(f"{path.name}: {str(e)}")
                    logging.exception(f"Erreur extraction PDF pour {path}")
                    
                done_files += 1
                after_func(0, lambda d=done_files: (
                    progress_bar.set(d / total_files),
                    progress_label.configure(text=f"Traité : {d} / {total_files} fichiers")
                ))

        except Exception as e:
            errors.append(str(e))
            logging.exception("Erreur globale dans PDF to Excel")
            
        def _final():
            cancelled = cancel_check_func()
            progress_bar.set(1 if not cancelled else min(done_files / total_files, 0.99))
            if errors:
                msg = "\n".join(errors[:5]) + ("\n..." if len(errors) > 5 else "")
                messagebox.showwarning("Terminé avec erreurs", f"Certaines extractions ont échoué :\n{msg}")
                
            status_txt = "annulée" if cancelled else "terminée"
            summary = f"Extraction {status_txt}. {done_files} fichier(s) traité(s)."
            progress_label.configure(text=f"✓ {summary}")
            stats_var.set(summary)
            log_func(f"[PDF to Excel] {summary}")
            
            if not cancelled:
                show_done_func("Extraction PDF vers Excel terminée", summary, str(out_dir))
                
            set_running_func(False)

        after_func(0, _final)

    threading.Thread(target=_worker, daemon=True).start()
