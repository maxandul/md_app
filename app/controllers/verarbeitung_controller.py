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
        
        # Eingangsordner ist IMMER ruecklauf/unverarbeitet (wo die eingehenden Dokumente liegen)
        # Korrektur: Von controllers/ aus 2 Ebenen hoch zur app/, dann Config-Pfad anwenden
        input_dir = Path(__file__).parent.parent / CFG["paths"]["ruecklauf"]["unverarbeitet"]
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
        
        # Zielpfad DS-Export: Jahr-spezifisch in tracking/ds_export/docx_extract_{jahr}.csv
        ds_export_cfg = (CFG.get("paths", {}) or {}).get("ds_export")
        # Korrektur: Von controllers/ aus 3 Ebenen hoch zur Root (md_app/)
        if ds_export_cfg:
            ds_export_path = Path(ds_export_cfg)
        else:
            # Default: tracking/ds_export/docx_extract_{jahr}.csv
            ds_export_path = Path(__file__).parent.parent.parent / "tracking" / "ds_export" / f"docx_extract_{durchlauf_jahr}.csv"
        
        export_ds_csv(
            results=all_results,
            out_csv=ds_export_path,
            sap_df=sap_df
        )
        
        # Verschiebe Dateien (nur Word-Dokumente, PDFs wurden bereits in process_pdfs verschoben)
        moved_ok, moved_man = move_after_processing(input_dir, docx_results)
        
        # Zähle alle Ergebnisse (Word + PDF)
        count_ok = sum(1 for r in all_results if r.get("status") == ProcStatus.OK.value)
        count_manuell = sum(1 for r in all_results if r.get("status") in (ProcStatus.MANUELL.value, ProcStatus.PRUEFUNG_NOETIG.value))
        
        # Status-Update
        app.proc_status.config(
            text=f"Verarbeitung abgeschlossen: {len(all_results)} Dateien, {count_ok} OK, {count_manuell} manuell",
            foreground="green"
        )
        logger.info("Verarbeitung abgeschlossen", extra={"count": len(all_results), "ok": count_ok, "manuell": count_manuell})
        
        # Treeview aktualisieren
        from views.ui_utils import autosize_tree_columns
        
        # Lösche alte Einträge
        for tree in [app.tree_proc, app.tree_pdfs]:
            for item in tree.get_children():
                tree.delete(item)
        
        # Füge neue Einträge hinzu - basierend auf Dateityp, nicht Status
        for result in all_results:
            fname = result.get("file", "")
            typ = result.get("typ", "")
            
            # Entscheide basierend auf Dateityp (Word vs. PDF)
            is_pdf = fname.lower().endswith(".pdf") or "PDF" in str(typ)
            
            if is_pdf:
                # PDF-Dokumente in PDF-Tabelle
                app.tree_pdfs.insert("", "end", values=[
                    result.get("file", ""),
                    result.get("typ", ""),
                    result.get("pn", ""),
                    result.get("name", ""),
                    result.get("status", ""),
                    result.get("reason", ""),
                    result.get("target", "")
                ])
            else:
                # Word-Dokumente in Word-Tabelle
                app.tree_proc.insert("", "end", values=[
                    result.get("file", ""),
                    result.get("typ", ""),
                    result.get("pn", ""),
                    result.get("name", ""),
                    result.get("status", ""),
                    result.get("reason", ""),
                    result.get("target", "")
                ])
        
        # Spaltenbreiten anpassen
        autosize_tree_columns(app.tree_proc)
        autosize_tree_columns(app.tree_pdfs)
        
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror(MDConstants.MSG_ERROR, f"Verarbeitung fehlgeschlagen: {e}")
        app.proc_status.config(text="Fehler bei der Verarbeitung", foreground="red")
        logger.error("Verarbeitung fehlgeschlagen", extra={"error": str(e)})