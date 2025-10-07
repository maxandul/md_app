from __future__ import annotations

from pathlib import Path


from app.constants import MDConstants, ProcStatus
from app.data_loader import load_employees
from app.services.document_service import process_docx_folder, process_pdfs
from app.services.export_service import export_sap_massenupload, export_ds_csv
from app.services.file_service import move_after_processing
from app.views.ui_utils import autosize_tree_columns
from app.logging_config import get_logger

logger = get_logger()


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
            messagebox.showerror(MDConstants.MSG_ERROR, f"Eingangsordner existiert nicht: {input_dir}")
            return
        
        # SAP-Daten laden
        from data_loader import load_employees
        sap_df = load_employees()
        
        # Durchlauf-Jahr
        durchlauf_jahr = app.proc_year_var.get()
        
        # Batch-Größe
        batch_size = app.batch_size_var.get()
        
        # Status-Update
        logger.info("Verarbeitung gestartet", extra={"input_dir": str(input_dir)})
        app.proc_status.config(text="Verarbeitung läuft...", foreground="blue")
        app.update_idletasks()
        
        # DOCX-Verarbeitung (nur Top-Level im projektweiten unverarbeitet-Ordner)
        docx_results = process_docx_folder(
            input_dir=input_dir,
            sap_df=sap_df,
            max_files=batch_size,
            durchlauf_jahr=durchlauf_jahr
        )
        
        # PDF-Verarbeitung
        # out_root: nutze config.paths.output_dir falls vorhanden, sonst UI-Wert (rpa_target_var)
        out_root_cfg = (CFG.get("paths", {}) or {}).get("output_dir")
        out_root_path = Path(out_root_cfg) if out_root_cfg else Path(app.rpa_target_var.get())

        # PDF-Verarbeitung ebenfalls aus demselben Top-Level-Ordner
        pdf_results = process_pdfs(
            in_dir=input_dir,
            out_root=out_root_path,
            sap_df=sap_df,
            durchlauf_jahr=durchlauf_jahr
        )
        
        # Kombiniere Ergebnisse
        all_results = docx_results + pdf_results
        
        # Export
        # Zielpfad SAP-Massenupload: config.paths.sap_massenupload, sonst Default in sap_massenupload/massenupload.xlsx
        sap_massenupload_cfg = (CFG.get("paths", {}) or {}).get("sap_massenupload")
        # Korrektur: Von controllers/ aus 3 Ebenen hoch zur Root (md_app/)
        sap_massenupload_path = Path(sap_massenupload_cfg) if sap_massenupload_cfg else (Path(__file__).parent.parent.parent / "sap_massenupload" / "massenupload.xlsx")
        export_sap_massenupload(
            results=all_results,
            sap_df=sap_df,
            out_xlsx=sap_massenupload_path
        )
        
        # Zielpfad DS-Export: config.paths.ds_export, sonst Default in tracking/ds_export/docx_extract.csv
        ds_export_cfg = (CFG.get("paths", {}) or {}).get("ds_export")
        # Korrektur: Von controllers/ aus 3 Ebenen hoch zur Root (md_app/)
        ds_export_path = Path(ds_export_cfg) if ds_export_cfg else (Path(__file__).parent.parent.parent / "tracking" / "ds_export" / "docx_extract.csv")
        export_ds_csv(
            results=all_results,
            out_csv=ds_export_path,
            sap_df=sap_df
        )
        
        # Verschiebe Dateien
        moved_ok, moved_man = move_after_processing(input_dir, all_results)
        
        # Status-Update
        app.proc_status.config(
            text=f"Verarbeitung abgeschlossen: {len(all_results)} Dateien, {moved_ok} OK, {moved_man} manuell",
            foreground="green"
        )
        logger.info("Verarbeitung abgeschlossen", extra={"count": len(all_results), "ok": moved_ok, "manuell": moved_man})
        
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
        messagebox.showerror(MDConstants.MSG_ERROR, f"Verarbeitung fehlgeschlagen: {e}")
        app.proc_status.config(text="Fehler bei der Verarbeitung", foreground="red")
        logger.error("Verarbeitung fehlgeschlagen", extra={"error": str(e)})