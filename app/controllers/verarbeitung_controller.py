from __future__ import annotations

from pathlib import Path

from tkinter import messagebox

from constants import MDConstants, ProcStatus
from data_loader import load_employees
from services.document_service import process_docx_folder, process_pdfs
from services.export_service import export_sap_massenupload, export_ds_csv
from services.file_service import move_after_processing
from views.ui_utils import autosize_tree_columns


def run_full_processing(app) -> None:
    """Führt die vollständige Dokumentenverarbeitung durch."""
    try:
        # Lade Konfiguration
        from data_loader import load_config
        CFG = load_config()
        
        # Eingangsordner
        input_dir = Path(app.rpa_target_var.get())
        if not input_dir.exists():
            from tkinter import messagebox
            messagebox.showerror("Fehler", f"Eingangsordner existiert nicht: {input_dir}")
            return
        
        # SAP-Daten laden
        from data_loader import load_employees
        sap_df = load_employees()
        
        # Durchlauf-Jahr
        durchlauf_jahr = app.proc_year_var.get()
        
        # Batch-Größe
        batch_size = app.batch_size_var.get()
        
        # Status-Update
        app.proc_status.config(text="Verarbeitung läuft...", foreground="blue")
        app.update_idletasks()
        
        # DOCX-Verarbeitung
        docx_results = process_docx_folder(
            input_dir=input_dir,
            sap_df=sap_df,
            max_files=batch_size,
            durchlauf_jahr=durchlauf_jahr
        )
        
        # PDF-Verarbeitung
        pdf_results = process_pdfs(
            in_dir=input_dir,
            out_root=Path(CFG["paths"]["output_dir"]),
            sap_df=sap_df,
            durchlauf_jahr=durchlauf_jahr
        )
        
        # Kombiniere Ergebnisse
        all_results = docx_results + pdf_results
        
        # Export
        export_sap_massenupload(
            results=all_results,
            sap_df=sap_df,
            out_xlsx=Path(CFG["paths"]["sap_massenupload"])
        )
        
        export_ds_csv(
            results=all_results,
            out_csv=Path(CFG["paths"]["ds_export"]),
            sap_df=sap_df
        )
        
        # Verschiebe Dateien
        moved_ok, moved_man = move_after_processing(input_dir, all_results)
        
        # Status-Update
        app.proc_status.config(
            text=f"Verarbeitung abgeschlossen: {len(all_results)} Dateien, {moved_ok} OK, {moved_man} manuell",
            foreground="green"
        )
        
        # Treeview aktualisieren
        from views.ui_utils import autosize_tree_columns
        
        # Lösche alte Einträge
        for tree in [app.tree_proc, app.tree_pdfs]:
            for item in tree.get_children():
                tree.delete(item)
        
        # Füge neue Einträge hinzu
        for result in all_results:
            if result.get("status") == ProcStatus.OK.value:
                app.tree_proc.insert("", "end", values=[
                    result.get("file", ""),
                    result.get("typ", ""),
                    result.get("pn", ""),
                    result.get("name", ""),
                    result.get("status", ""),
                    result.get("reason", "")
                ])
            else:
                app.tree_pdfs.insert("", "end", values=[
                    result.get("file", ""),
                    result.get("typ", ""),
                    result.get("pn", ""),
                    result.get("name", ""),
                    result.get("status", ""),
                    result.get("reason", "")
                ])
        
        # Spaltenbreiten anpassen
        autosize_tree_columns(app.tree_proc)
        autosize_tree_columns(app.tree_pdfs)
        
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("Fehler", f"Verarbeitung fehlgeschlagen: {e}")
        app.proc_status.config(text="Fehler bei der Verarbeitung", foreground="red")