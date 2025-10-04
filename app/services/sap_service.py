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

from constants import MDConstants, ProcStatus, DashTag
from data_loader import load_employees, load_config


def check_stammdaten(app) -> None:
    """Controller: Prüft EXPORT.xlsx, befüllt Prüftabellen und Label im UI."""
    from datetime import datetime
    CFG = load_config()
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
        app.tree_checks.insert("", "end", values=["Datei laden", f"Fehler: {e}"])
        return

    # 4. Pflichtspalten-Validierung
    required_cols = MDConstants.REQUIRED_COLS
    df_cols = set(df.columns)
    missing = [c for c in required_cols if c not in df_cols]
    if missing:
        app.tree_checks.insert("", "end", values=["Pflichtspalten vorhanden", f"FEHLT: {', '.join(missing)}"])
    else:
        app.tree_checks.insert("", "end", values=["Pflichtspalten vorhanden", "OK"])

    # 5. Beschäftigungsgrad = 0 prüfen
    if "BsGrd" in df.columns:
        bs0 = df[df["BsGrd"].astype(str).str.strip().isin(["0","0.0"])]
        for _, r in bs0.iterrows():
            app.tree_findings.insert("", "end", values=[
                "BsGrd=0",
                str(r.get("ID_NO_ZERO","")),
                str(r.get("Nachname","")),
                str(r.get("Rufname","")),
                "Beschäftigungsgrad=0"
            ])

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
            ])

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
            ])

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
        from data_loader import load_config
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


def refresh_dashboard(app) -> None:
    """Lädt Dashboard-Daten gemäß aktueller Filter und befüllt den Treeview."""
    for item in app.tree_dashboard.get_children():
        app.tree_dashboard.delete(item)

    try:
        name_search = app.dash_name_search.get().strip().lower()
        status_filter = app.dash_status_filter.get().strip()

        df = app.tracking.get_dashboard_data(filter_status=status_filter)
        if df.empty:
            return

        if name_search:
            name_mask = (
                df["vg_name"].astype(str).str.lower().str.contains(name_search, na=False)
                | df["ma_name"].astype(str).str.lower().str.contains(name_search, na=False)
            )
            df = df[name_mask]

        def safe_value(val):
            try:
                if pd.isna(val) or val == "nan" or val == "NaN":
                    return ""
            except Exception:
                pass
            return str(val) if val is not None else ""

        def format_pn(val):
            try:
                if isinstance(val, float) and float(val).is_integer():
                    return str(int(val))
                s = safe_value(val)
                if s.endswith('.0') and s.replace('.', '', 1).isdigit():
                    return s[:-2]
                return s
            except Exception:
                return safe_value(val)

        def status_tag_for(status: str) -> str:
            s = (status or "").lower().strip()
            if s == ProcStatus.AUSSTEHEND.value:
                return DashTag.AUSSTEHEND.value
            if s == ProcStatus.ERHALTEN.value:
                return DashTag.ERHALTEN.value
            if s == ProcStatus.ERUEBRIGT.value:
                return DashTag.ERUEBRIGT.value
            if s == ProcStatus.PRUEFUNG_NOETIG.value:
                return DashTag.PRUEFUNG_NOETIG.value
            if s == ProcStatus.OK.value:
                return DashTag.OK.value
            if s == ProcStatus.MANUELL.value:
                return DashTag.MANUELL.value
            return ""

        for _, row in df.iterrows():
            status = safe_value(row.get("status", ""))
            tag = status_tag_for(status)
            app.tree_dashboard.insert(
                "",
                "end",
                values=(
                    safe_value(row.get("log_id", "")),
                    format_pn(row.get("vg_pn", "")),
                    safe_value(row.get("vg_name", "")),
                    format_pn(row.get("ma_pn", "")),
                    safe_value(row.get("ma_name", "")),
                    safe_value(row.get("doc_type", "")),
                    safe_value(row.get("erwartet", "")),
                    safe_value(row.get("erhalten", "")),
                    status,
                    safe_value(row.get("status_grund", "")),
                    safe_value(row.get("versendet_am", "")),
                    safe_value(row.get("zuletzt_erinnert_am", "")),
                ),
                tags=(tag,) if tag else (),
            )
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Fehler beim Laden der Dashboard-Daten: {e}")


def export_dashboard(app) -> None:
    """Exportiert die aktuell gefilterten Dashboard-Daten als CSV."""
    try:
        from tkinter import filedialog

        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not filename:
            return

        status_filter = app.dash_status_filter.get().strip()
        df = app.tracking.get_dashboard_data(filter_status=status_filter)
        df.to_csv(filename, sep=";", index=False, encoding="utf-8-sig")
        messagebox.showinfo(MDConstants.MSG_SUCCESS, f"Dashboard-Daten exportiert nach: {filename}")
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Export fehlgeschlagen: {e}")


def manual_adjustment(app) -> None:
    """Öffnet einen Dialog zur manuellen Anpassung eines ausgewählten Dashboard-Eintrags."""
    selection = app.tree_dashboard.selection()
    if not selection:
        messagebox.showwarning(MDConstants.MSG_WARNING, "Bitte wählen Sie einen Eintrag aus.")
        return

    item = app.tree_dashboard.item(selection[0])
    values = item.get("values", [])
    if len(values) < 12:
        messagebox.showerror(MDConstants.MSG_ERROR, "Ungültiger Eintrag ausgewählt.")
        return

    log_id = values[0]

    dialog = tk.Toplevel(app)
    dialog.title("Manuelle Anpassung")
    dialog.geometry("700x600")
    dialog.transient(app)
    dialog.grab_set()

    canvas = tk.Canvas(dialog)
    scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    ttk.Label(scrollable_frame, text=f"Log-ID: {log_id}", font=("Arial", 12, "bold")).pack(pady=8)

    columns = [
        ("vg_pn", "VG PN", True),
        ("vg_name", "VG Name", True),
        ("ma_pn", "MA PN", True),
        ("ma_name", "MA Name", True),
        ("doc_type", "Dokument-Typ", False),
        ("erwartet", "Erwartet", True),
        ("erhalten", "Erhalten", True),
        ("status", "Status", True),
        ("status_grund", "Status Grund", True),
        ("versendet_am", "Versendet am", True),
        ("zuletzt_erinnert_am", "Zuletzt erinnert am", True),
    ]

    entry_vars: dict[str, tk.StringVar] = {}

    for i, (col_key, col_label, editable) in enumerate(columns):
        frame = ttk.Frame(scrollable_frame)
        frame.pack(fill="x", padx=8, pady=2)

        ttk.Label(frame, text=f"{col_label}:", width=20, anchor="w").pack(side="left")

        if not editable:
            ttk.Label(frame, text=str(values[i + 1]), relief="sunken", width=40).pack(side="left", padx=(4, 0))
        else:
            if col_key == "status":
                var = tk.StringVar(value=values[i + 1])
                entry = ttk.Combobox(
                    frame,
                    textvariable=var,
                    width=37,
                    values=[
                        ProcStatus.AUSSTEHEND.value,
                        ProcStatus.ERHALTEN.value,
                        ProcStatus.PRUEFUNG_NOETIG.value,
                        ProcStatus.ERUEBRIGT.value,
                    ],
                )
            elif col_key == "status_grund":
                var = tk.StringVar(value=values[i + 1])
                entry = ttk.Combobox(
                    frame,
                    textvariable=var,
                    width=37,
                    values=[
                        "",
                        "Grund_Prüfung (aus Verarbeitung)",
                        "Krankheit/Unfall",
                        "anderer VG",
                        "Austritt",
                        "sonstiges",
                    ],
                )
            else:
                var = tk.StringVar(value=str(values[i + 1]))
                entry = ttk.Entry(frame, textvariable=var, width=40)

            entry.pack(side="left", padx=(4, 0))
            entry_vars[col_key] = var

    button_frame = ttk.Frame(scrollable_frame)
    button_frame.pack(pady=16)

    def apply_adjustment() -> None:
        try:
            updates: dict[str, str] = {}
            for col_key, var in entry_vars.items():
                new_value = var.get().strip()
                col_index = None
                for idx, (key, _, _) in enumerate(columns):
                    if key == col_key:
                        col_index = idx + 1
                        break
                if col_index is not None:
                    old_value = str(values[col_index])
                    if new_value != old_value:
                        updates[col_key] = new_value

            if not updates:
                messagebox.showinfo(MDConstants.MSG_INFO, "Keine Änderungen vorgenommen.")
                dialog.destroy()
                return

            if app.tracking.update_entry(log_id, updates):
                messagebox.showinfo(MDConstants.MSG_SUCCESS, f"Anpassung gespeichert.\nGeändert: {', '.join(updates.keys())}")
                dialog.destroy()
                refresh_dashboard(app)
            else:
                messagebox.showerror(MDConstants.MSG_ERROR, "Anpassung fehlgeschlagen - Eintrag nicht gefunden.")

        except Exception as e:
            messagebox.showerror(MDConstants.MSG_ERROR, f"Fehler beim Speichern: {e}")

    ttk.Button(button_frame, text="Anwenden", command=apply_adjustment).pack(side="left", padx=(0, 8))
    ttk.Button(button_frame, text="Abbrechen", command=dialog.destroy).pack(side="left")

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

def refresh_mgr_table(app) -> None:
    """Aktualisiert die VG-Tabelle im Massenversand-Tab."""
    # Treeview leeren
    for item in app.tree.get_children():
        app.tree.delete(item)

    # Filter anwenden
    filter_text = app.filter_var.get().lower() if hasattr(app, 'filter_var') else ""

    for vg_pn, pack in app.mgr_index.items():
        mgr = pack["manager"]
        if mgr is None:
            continue

        # Filter prüfen
        if filter_text:
            mgr_text = f"{mgr.get('Nachname','')} {mgr.get('Rufname','')} {mgr.get('OE Kurzb.','')}".lower()
            if filter_text not in mgr_text:
                continue

        subs_count = len(pack["subs"])
        app.tree.insert("", "end", iid=vg_pn, values=[
            vg_pn,
            mgr.get("Nachname", ""),
            mgr.get("Rufname", ""),
            mgr.get("OE Kurzb.", ""),
            subs_count,
        ])


def refresh_mgr_table_einzel(app) -> None:
    """Aktualisiert die VG-Tabelle im Einzelversand-Tab."""
    # Treeview leeren
    for item in app.tree_einzel.get_children():
        app.tree_einzel.delete(item)

    filter_text = app.filter_var.get().lower() if hasattr(app, 'filter_var') else ""

    for vg_pn, pack in app.mgr_index.items():
        mgr = pack["manager"]
        if mgr is None:
            continue

        if filter_text:
            mgr_text = f"{mgr.get('Nachname','')} {mgr.get('Rufname','')} {mgr.get('OE Kurzb.','')}".lower()
            if filter_text not in mgr_text:
                continue

        subs_count = len(pack["subs"])
        app.tree_einzel.insert("", "end", iid=vg_pn, values=[
            vg_pn,
            mgr.get("Nachname", ""),
            mgr.get("Rufname", ""),
            mgr.get("OE Kurzb.", ""),
            subs_count,
        ])


def refresh_vg_list(app) -> None:
    """Aktualisiert die VG-Liste im VG-MA-Tab."""
    for item in app.vg_tree.get_children():
        app.vg_tree.delete(item)

    filter_text = app.vg_search_var.get().lower() if hasattr(app, 'vg_search_var') else ""
    for vg_pn, pack in app.mgr_index.items():
        mgr = pack["manager"]
        if mgr is None:
            continue
        if filter_text:
            mgr_text = f"{mgr.get('Nachname','')} {mgr.get('Rufname','')} {mgr.get('OE Kurzb.','')}".lower()
            if filter_text not in mgr_text:
                continue
        app.vg_tree.insert("", "end", iid=vg_pn, values=[
            vg_pn,
            mgr.get("Nachname", ""),
            mgr.get("Rufname", ""),
            mgr.get("OE Kurzb.", ""),
        ])


def refresh_ma_list(app) -> None:
    """Aktualisiert die MA-Liste im VG-MA-Tab."""
    for item in app.ma_tree.get_children():
        app.ma_tree.delete(item)

    filter_text = app.ma_search_var.get().lower() if hasattr(app, 'ma_search_var') else ""

    for vg_pn, pack in app.mgr_index.items():
        subs = pack["subs"]
        for _, r in subs.iterrows():
            if filter_text:
                txt = f"{r.get('Nachname','')} {r.get('Rufname','')} {r.get('OE Kurzb.','')}".lower()
                if filter_text not in txt:
                    continue
            app.ma_tree.insert("", "end", iid=str(r.get("ID_NO_ZERO","")), values=[
                str(r.get("ID_NO_ZERO","")),
                str(r.get("Nachname","")),
                str(r.get("Rufname","")),
                str(r.get("OE Kurzb.","")),
            ])


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