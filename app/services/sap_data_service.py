"""
SAP-Service für das MD-Prozess-Tool.

Dieser Service kapselt alle Funktionen zur SAP-Datenverarbeitung,
einschließlich Stammdaten-Validierung, Dashboard-Management und
Organisationsstruktur-Handling.
"""

from __future__ import annotations
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime

from app.constants import MDConstants, ProcStatus, DashTag
from app.data_loader import load_employees, load_config
from app.theme import get_row_tag


def check_stammdaten(app) -> None:
    """Controller: Prüft EXPORT.xlsx, befüllt Prüftabellen und Label im UI."""
    from datetime import datetime
    CFG = load_config()
    # Korrektur: Von services/ aus 2 Ebenen hoch zur app/, Config-Pfade sind relativ zu app/
    xlsx_path = Path(__file__).parent.parent / CFG["paths"]["sap_stammdaten"]

    # 1. Dateiinfo und Metadaten prüfen
    try:
        mtime = datetime.fromtimestamp((Path(xlsx_path)).stat().st_mtime)
        app.lbl_fileinfo.config(text=f"Datei: {xlsx_path} • Letzte Änderung: {mtime:%Y-%m-%d %H:%M}", foreground="black")
    except Exception as e:
        app.lbl_fileinfo.config(text=f"Datei: {xlsx_path} • Fehler beim Lesen: {e}", foreground="red")

    # 2. Vorherige Ergebnisse löschen
    for t in (app.tree_checks, app.tree_findings):
        for i in t.get_children():
            t.delete(i)

    # 3. Excel-Datei laden
    try:
        df = load_employees()
    except Exception as e:
        app.tree_checks.insert("", "end", values=["Datei laden", f"Fehler: {e}"], tags=(get_row_tag(0),))
        return

    # 4. Pflichtspalten-Validierung
    required_cols = MDConstants.REQUIRED_COLS
    df_cols = set(df.columns)
    missing = [c for c in required_cols if c not in df_cols]
    if missing:
        app.tree_checks.insert("", "end", values=["Pflichtspalten vorhanden", f"FEHLT: {', '.join(missing)}"], tags=(get_row_tag(0),))
    else:
        app.tree_checks.insert("", "end", values=["Pflichtspalten vorhanden", "OK"], tags=(get_row_tag(0),))

    # 5. Beschäftigungsgrad = 0 prüfen
    findings_idx = 0
    if "BsGrd" in df.columns:
        bs0 = df[df["BsGrd"].astype(str).str.strip().isin(["0","0.0"])]
        for _, r in bs0.iterrows():
            app.tree_findings.insert("", "end", values=[
                "BsGrd=0",
                str(r.get("ID_NO_ZERO","")),
                str(r.get("Nachname","")),
                str(r.get("Rufname","")),
                "Beschäftigungsgrad=0"
            ], tags=(get_row_tag(findings_idx),))
            findings_idx += 1

    # 6. Duplikat-Erkennung bei Personalnummern
    if "ID_NO_ZERO" in df.columns:
        pns = df["ID_NO_ZERO"].astype(str).str.strip()
        dup_mask = pns.duplicated(keep=False)
        dups = df[dup_mask].copy()
        dups = dups.assign(_pn=pns[dup_mask]).sort_values("_pn")
        for _, r in dups.iterrows():
            app.tree_findings.insert("", "end", values=[
                "Duplikat PN",
                str(r.get("ID_NO_ZERO","")),
                str(r.get("Nachname","")),
                str(r.get("Rufname","")),
                "Mehrfach vorhandene Personalnummer"
            ], tags=(get_row_tag(findings_idx),))
            findings_idx += 1

    # 7. Vorgesetzten-PN Validierung (leer oder nur Nullen)
    if "Dir. Vorgesetzter (PN)" in df.columns:
        vg_col = df["Dir. Vorgesetzter (PN)"].astype(str).str.strip()
        bad_vg = vg_col.isin(["", "nan", "NaN", "None"]) | vg_col.str.fullmatch(r"0+")
        vg0 = df[bad_vg]
        for _, r in vg0.iterrows():
            app.tree_findings.insert("", "end", values=[
                "VG_PN fehlend/0",
                str(r.get("ID_NO_ZERO","")),
                str(r.get("Nachname","")),
                str(r.get("Rufname","")),
                "Kein gültiger Wert in 'Dir. Vorgesetzter (PN)'"
            ], tags=(get_row_tag(findings_idx),))
            findings_idx += 1

    try:
        from views.ui_utils import autosize_tree_columns
        autosize_tree_columns(app.tree_checks)
        autosize_tree_columns(app.tree_findings)
    except Exception:
        pass


def create_vg_ma_relationship(app) -> None:
    """Erstellt neue VG-MA-Beziehung in EXPORT.xlsx."""
    vg_sel = app.vg_tree.selection()
    ma_sel = app.ma_tree.selection()
    
    if not vg_sel or not ma_sel:
        messagebox.showwarning(MDConstants.MSG_WARNING, "Bitte wählen Sie sowohl einen Vorgesetzten als auch einen Mitarbeitenden aus.")
        return
    
    vg_item = app.vg_tree.item(vg_sel[0])
    ma_item = app.ma_tree.item(ma_sel[0])
    vg_pn = str(vg_item['values'][0])  # Explizit als String konvertieren
    ma_pn = str(ma_item['values'][0])   # Explizit als String konvertieren
    
    try:
        # Lade aktuelle EXPORT.xlsx mit gleicher Normalisierung wie load_employees()
        from app.data_loader import load_config
        CFG = load_config()
        xlsx_path = Path(__file__).parent.parent / CFG["paths"]["sap_stammdaten"]
        df = pd.read_excel(xlsx_path)
        
        # Gleiche Datentyp-Konvertierung wie in load_employees()
        for col, t in {"ID_NO_ZERO": str, "Dir. Vorgesetzter (PN)": str, "Ans.": str}.items():
            if col in df.columns:
                df[col] = df[col].astype(t)
        df.columns = [c.strip() for c in df.columns]
        
        # Personalnummern normalisieren (gleiche Logik wie in load_employees)
        if "ID_NO_ZERO" in df.columns:
            df["ID_NO_ZERO"] = df["ID_NO_ZERO"].astype(str).str.strip()
        if "Dir. Vorgesetzter (PN)" in df.columns:
            df["Dir. Vorgesetzter (PN)"] = df["Dir. Vorgesetzter (PN)"].astype(str).str.strip()
        
        # IDs mit führenden Nullen angleichen
        if "ID_NO_ZERO" in df.columns and "Dir. Vorgesetzter (PN)" in df.columns:
            max_len = max(
                df["ID_NO_ZERO"].str.len().max(),
                df["Dir. Vorgesetzter (PN)"].str.len().max()
            )
            df["ID_NO_ZERO"] = df["ID_NO_ZERO"].str.zfill(max_len)
            df["Dir. Vorgesetzter (PN)"] = df["Dir. Vorgesetzter (PN)"].str.zfill(max_len)
        
        # Finde MA-Zeile mit normalisierter PN
        ma_row = df[df["ID_NO_ZERO"].astype(str) == ma_pn]
        if ma_row.empty:
            # Debug: Zeige verfügbare PNs und vergleiche mit app.df
            available_pns = df["ID_NO_ZERO"].astype(str).tolist()[:10]  # Erste 10 PNs aus Excel
            loaded_pns = app.df["ID_NO_ZERO"].astype(str).tolist()[:10]  # Erste 10 PNs aus geladenem DataFrame
            
            # Prüfe ob PN in geladenem DataFrame existiert
            ma_in_loaded = app.df[app.df["ID_NO_ZERO"].astype(str) == ma_pn]
            
            debug_msg = f"Mitarbeiter mit PN {ma_pn} nicht in EXPORT.xlsx gefunden.\n\n"
            debug_msg += f"Verfügbare PNs in Excel (erste 10): {available_pns}\n"
            debug_msg += f"Verfügbare PNs in geladenem DF (erste 10): {loaded_pns}\n"
            debug_msg += f"Gesuchte PN: '{ma_pn}' (Typ: {type(ma_pn)})\n"
            debug_msg += f"PN in geladenem DF gefunden: {not ma_in_loaded.empty}\n"
            debug_msg += f"Excel-Datei Zeilen: {len(df)}, Geladener DF Zeilen: {len(app.df)}"
            
            messagebox.showerror(MDConstants.MSG_ERROR, debug_msg)
            return
        
        # Erstelle Kopie der MA-Zeile
        new_row = ma_row.iloc[0].copy()
        new_row["Dir. Vorgesetzter (PN)"] = vg_pn
        
        # Füge neue Zeile hinzu
        df = pd.concat([df, new_row.to_frame().T], ignore_index=True)
        
        # Speichere zurück
        df.to_excel(xlsx_path, index=False)
        
        messagebox.showinfo("Erfolg", 
            f"Neue VG-MA-Beziehung erstellt:\n"
            f"VG: {vg_item['values'][2]} {vg_item['values'][1]} ({vg_pn})\n"
            f"MA: {ma_item['values'][2]} {ma_item['values'][1]} ({ma_pn})\n\n"
            f"Bitte starten Sie die App neu, um die Änderungen zu laden."
        )
        
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Erstellen der Beziehung: {e}")

def refresh_mgr_table(app) -> None:
    """Aktualisiert die VG-Tabelle im Massenversand-Tab."""
    # Treeview leeren
    for item in app.tree.get_children():
        app.tree.delete(item)

    # Filter anwenden
    filter_text = app.filter_var.get().lower() if hasattr(app, 'filter_var') else ""
    nur_nicht_versendet = app.nur_nicht_versendet_var.get() if hasattr(app, 'nur_nicht_versendet_var') else False
    
    # Tracking-Daten laden um versendete VGs zu identifizieren
    jahr = app.md_durchlauf_jahr.get()
    tracking_path = Path(__file__).parent.parent / f"../tracking/md_logging_{jahr}.csv"
    
    vg_dokument_count = {}  # VG-PN -> Anzahl Dokumente im Tracking
    if tracking_path.exists():
        try:
            df_tracking = pd.read_csv(tracking_path, sep=";", encoding="utf-8-sig")
            # Normalisiere VG-PNs (wichtig für Vergleich!)
            if "vg_pn" in df_tracking.columns:
                df_tracking["vg_pn"] = df_tracking["vg_pn"].astype(str).str.strip()
                
                # Entferne .0 falls vorhanden (Excel-Import-Artefakt)
                df_tracking["vg_pn"] = df_tracking["vg_pn"].str.replace(r'\.0$', '', regex=True)
                
                # Zähle Dokumente pro VG
                vg_counts = df_tracking.groupby("vg_pn").size()
                vg_dokument_count = vg_counts.to_dict()
                
                # Debug-Ausgabe in Log
                from app.logging_config import get_logger
                logger = get_logger()
                logger.info(f"Tracking geladen für Jahr {jahr}: {len(vg_dokument_count)} VGs mit Einträgen")
        except Exception as e:
            from app.logging_config import get_logger
            logger = get_logger()
            logger.warning(f"Fehler beim Laden von Tracking für Jahr {jahr}: {e}")

    # Gefilterte Items sammeln
    filtered_items = []
    for vg_pn, pack in app.mgr_index.items():
        mgr = pack["manager"]
        if mgr is None:
            continue

        # Versand-Status prüfen (≥3 Dokumente = Massenversand)
        # Normalisiere VG-PN für Lookup (entferne .0 falls vorhanden)
        vg_pn_normalized = str(vg_pn).strip()
        if vg_pn_normalized.endswith('.0'):
            vg_pn_normalized = vg_pn_normalized[:-2]
        
        anzahl_dokumente = vg_dokument_count.get(vg_pn_normalized, 0)
        ist_versendet = anzahl_dokumente >= 3
        
        # Debug für erste 3 VGs
        if len(filtered_items) < 3:
            from app.logging_config import get_logger
            logger = get_logger()
            logger.info(f"VG {vg_pn_normalized}: {anzahl_dokumente} Dokumente, versendet={ist_versendet}")
        
        # Filter: Nur nicht versendete
        if nur_nicht_versendet and ist_versendet:
            continue

        # Filter: Suchtext
        if filter_text:
            mgr_text = f"{mgr.get('Nachname','')} {mgr.get('Rufname','')} {mgr.get('OE Kurzb.','')}".lower()
            if filter_text not in mgr_text:
                continue

        filtered_items.append((vg_pn, pack, ist_versendet))

    # Items mit eindeutigen IDs einfügen (konsistent mit versand_view.py)
    for i, (vg_pn, pack, ist_versendet) in enumerate(filtered_items):
        mgr = pack["manager"]
        subs_count = len(pack["subs"])
        # iid darf beliebig sein; PN immer aus values[0] lesen!
        unique_id = f"{vg_pn}_{i}"
        
        # Tag basierend auf Versand-Status
        if ist_versendet:
            tag = "versendet"
            # Debug: Zeige welche VGs grün markiert werden
            if i < 5:  # Nur erste 5 loggen
                from app.logging_config import get_logger
                logger = get_logger()
                logger.info(f"VG {mgr.get('Nachname','')} ({vg_pn_normalized}) wird GRÜN markiert ({anzahl_dokumente} Dokumente)")
        else:
            tag = get_row_tag(i)
        
        app.tree.insert("", "end", iid=unique_id, values=[
            vg_pn,
            mgr.get("Nachname", ""),
            mgr.get("Rufname", ""),
            mgr.get("OE Kurzb.", ""),
            subs_count,
        ], tags=(tag,))
    
    # Automatische Spaltenbreiten-Anpassung
    try:
        from app.views.ui_utils import autosize_tree_columns
        autosize_tree_columns(app.tree)
    except Exception:
        pass


def refresh_mgr_table_einzel(app) -> None:
    """Aktualisiert die VG-Tabelle im Einzelversand-Tab."""
    # Treeview leeren
    for item in app.tree_einzel.get_children():
        app.tree_einzel.delete(item)

    # Verwende filter_var_einzel für Einzelversand-Tab
    filter_text = app.filter_var_einzel.get().lower() if hasattr(app, 'filter_var_einzel') else ""

    # Gefilterte Items sammeln
    filtered_items = []
    for vg_pn, pack in app.mgr_index.items():
        mgr = pack["manager"]
        if mgr is None:
            continue

        # Filter prüfen (analog zum Massenversand)
        if filter_text:
            mgr_text = f"{mgr.get('Nachname','')} {mgr.get('Rufname','')} {mgr.get('OE Kurzb.','')} {vg_pn}".lower()
            if filter_text not in mgr_text:
                continue

        filtered_items.append((vg_pn, pack))

    # Items mit eindeutigen IDs einfügen (konsistent mit versand_view.py)
    for i, (vg_pn, pack) in enumerate(filtered_items):
        mgr = pack["manager"]
        subs_count = len(pack["subs"])
        # iid darf beliebig sein; PN immer aus values[0] lesen!
        unique_id = f"{vg_pn}_{i}"
        app.tree_einzel.insert("", "end", iid=unique_id, values=[
            vg_pn,
            mgr.get("Nachname", ""),
            mgr.get("Rufname", ""),
            mgr.get("OE Kurzb.", ""),
            subs_count,
        ], tags=(get_row_tag(i),))
    
    # Automatische Spaltenbreiten-Anpassung
    try:
        from app.views.ui_utils import autosize_tree_columns
        autosize_tree_columns(app.tree_einzel)
    except Exception:
        pass


def refresh_vg_list(app) -> None:
    """Aktualisiert die VG-Liste im VG-MA-Tab."""
    for item in app.vg_tree.get_children():
        app.vg_tree.delete(item)

    filter_text = app.vg_search_var.get().lower() if hasattr(app, 'vg_search_var') else ""
    
    # Gefilterte Items sammeln
    filtered_items = []
    for vg_pn, pack in app.mgr_index.items():
        mgr = pack["manager"]
        if mgr is None:
            continue
        if filter_text:
            mgr_text = f"{mgr.get('Nachname','')} {mgr.get('Rufname','')} {mgr.get('OE Kurzb.','')}".lower()
            if filter_text not in mgr_text:
                continue
        
        filtered_items.append((vg_pn, mgr))
    
    # Items mit eindeutigen IDs einfügen (konsistent mit versand_view.py)
    for i, (vg_pn, mgr) in enumerate(filtered_items):
        unique_id = f"{vg_pn}_{i}"
        app.vg_tree.insert("", "end", iid=unique_id, values=[
            vg_pn,
            mgr.get("Nachname", ""),
            mgr.get("Rufname", ""),
            mgr.get("OE Kurzb.", ""),
        ], tags=(get_row_tag(i),))
    
    # Automatische Spaltenbreiten-Anpassung
    try:
        from app.views.ui_utils import autosize_tree_columns
        autosize_tree_columns(app.vg_tree)
    except Exception:
        pass


def refresh_ma_list(app) -> None:
    """Aktualisiert die MA-Liste im VG-MA-Tab."""
    for item in app.ma_tree.get_children():
        app.ma_tree.delete(item)

    filter_text = app.ma_search_var.get().lower() if hasattr(app, 'ma_search_var') else ""

    # Gefilterte Items sammeln
    filtered_items = []
    for vg_pn, pack in app.mgr_index.items():
        subs = pack["subs"]
        for _, r in subs.iterrows():
            if filter_text:
                txt = f"{r.get('Nachname','')} {r.get('Rufname','')} {r.get('OE Kurzb.','')}".lower()
                if filter_text not in txt:
                    continue
            
            filtered_items.append(r)

    # Items mit eindeutigen IDs einfügen (konsistent mit versand_view.py)
    for i, r in enumerate(filtered_items):
        ma_pn = str(r.get("ID_NO_ZERO",""))
        unique_id = f"{ma_pn}_{i}"
        app.ma_tree.insert("", "end", iid=unique_id, values=[
            ma_pn,
            str(r.get("Nachname","")),
            str(r.get("Rufname","")),
            str(r.get("OE Kurzb.","")),
        ], tags=(get_row_tag(i),))
    
    # Automatische Spaltenbreiten-Anpassung
    try:
        from app.views.ui_utils import autosize_tree_columns
        autosize_tree_columns(app.ma_tree)
    except Exception:
        pass


def update_selection_status(app) -> None:
    """Aktualisiert den Auswahlinfo-Text im VG-MA-Tab."""
    vg_sel = app.vg_tree.selection() if hasattr(app, 'vg_tree') else []
    ma_sel = app.ma_tree.selection() if hasattr(app, 'ma_tree') else []
    if vg_sel and ma_sel:
        vg = vg_sel[0]
        ma = ma_sel[0]
        app.selection_status.config(text=f"VG: {vg} • MA: {ma}")
    elif vg_sel:
        app.selection_status.config(text=f"VG: {vg_sel[0]} • MA: -")
    elif ma_sel:
        app.selection_status.config(text=f"VG: - • MA: {ma_sel[0]}")
    else:
        app.selection_status.config(text="Keine Auswahl")