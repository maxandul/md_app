"""
Mail-Service für das MD-Prozess-Tool.

Dieser Service kapselt alle Funktionen zur E-Mail-Verarbeitung,
einschließlich Outlook-Integration, Anhang-Verarbeitung und
Mail-Versand-Funktionalitäten.
"""

from __future__ import annotations
from pathlib import Path
import os
import shutil
import win32com.client
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import date
import pandas as pd

from app.constants import MDConstants, ProcStatus
from app.logging_config import get_logger
logger = get_logger()
from app.services.dispatch_service import build_and_send_for_manager
from app.data_loader import load_employees, build_manager_index, load_config
from app.views.ui_utils import autosize_tree_columns

def send_managers(app, mode: str | None = None) -> None:
    """Sendet MD-Dokumente an ausgewählte Vorgesetzte."""
    selected_items = app.tree.selection()
    if not selected_items:
        raise ValueError("Bitte wählen Sie mindestens einen Vorgesetzten aus.")
    
    try:
        rb_year = app.rb_year_var.get()
        ab_year = app.ab_year_var.get()
        today = date.today()
        
        # Lade Konfiguration
        CFG = load_config()
        out_root = Path(CFG["paths"]["output_dir"])
        
        sent_count = 0
        logger.info("Versand gestartet", extra={"count_selected": len(selected_items)})
        for item_id in selected_items:
            # PN aus Spaltenwerten ermitteln (iid ist z.B. "<PN>_<index>")
            try:
                values = app.tree.item(item_id, "values")
                vg_pn = str(values[0]) if values else ""
            except Exception:
                vg_pn = ""
            if vg_pn not in app.mgr_index:
                continue
                
            pack = app.mgr_index[vg_pn]
            mgr = pack["manager"]
            subs = pack["subs"]
            
            if mgr is None or subs.empty:
                continue
                
            try:
                build_and_send_for_manager(
                    mgr_row=mgr,
                    subs_df=subs,
                    rb_year=rb_year,
                    ab_year=ab_year,
                    today=today,
                    out_root=out_root,
                    managers_index=app.mgr_index,
                    include_feedback=app.var_rb.get(),
                    send_mode=mode
                )
                sent_count += 1
            except Exception as e:
                logger.error("Versandfehler", extra={"vg_pn": vg_pn, "error": str(e)})
                raise RuntimeError(f"Fehler beim Versand an {mgr.get('Nachname', '')} {mgr.get('Rufname', '')}:") from e
        
        logger.info("Versand abgeschlossen", extra={"sent_count": sent_count})
        # Erfolgsmeldung im Controller anzeigen
        
    except Exception as e:
        raise


def send_selected_employees(app, mode: str | None = None) -> None:
    """Sendet MD-Dokumente an ausgewählte Mitarbeiter."""
    selected_items = app.tree_einzel.selection()
    if not selected_items:
        raise ValueError("Bitte wählen Sie mindestens einen Mitarbeiter aus.")
    
    try:
        rb_year = app.rb_year_var_einzel.get()
        ab_year = app.ab_year_var_einzel.get()
        today = date.today()
        
        # Lade Konfiguration
        CFG = load_config()
        out_root = Path(CFG["paths"]["output_dir"])
        
        sent_count = 0
        logger.info("Einzelversand gestartet", extra={"count_selected": len(selected_items)})
        for item_id in selected_items:
            # PN aus Spaltenwerten ermitteln (iid ist z.B. "<PN>_<index>")
            try:
                values = app.tree_einzel.item(item_id, "values")
                vg_pn = str(values[0]) if values else ""
            except Exception:
                vg_pn = ""
            if vg_pn not in app.mgr_index:
                continue
                
            pack = app.mgr_index[vg_pn]
            mgr = pack["manager"]
            subs = pack["subs"]
            
            if mgr is None or subs.empty:
                continue
                
            try:
                build_and_send_for_manager(
                    mgr_row=mgr,
                    subs_df=subs,
                    rb_year=rb_year,
                    ab_year=ab_year,
                    today=today,
                    out_root=out_root,
                    managers_index=app.mgr_index,
                    include_feedback=app.var_rb.get(),
                    send_mode=mode
                )
                sent_count += 1
            except Exception as e:
                raise RuntimeError(f"Fehler beim Versand an {mgr.get('Nachname', '')} {mgr.get('Rufname', '')}:") from e
        
        # Erfolgsmeldung im Controller anzeigen
        
    except Exception as e:
        raise


def render_mail_preview(app) -> None:
    """Rendert eine Vorschau der E-Mail-Inhalte."""
    # Diese Funktion kann erweitert werden, um eine detaillierte E-Mail-Vorschau zu zeigen
    # Info im Controller anzeigen

    if hasattr(app, "inbox_status"):
        app.inbox_status.config(text="Bereit.", foreground="gray")

    try:
        autosize_tree_columns(app.tree_ok)
        autosize_tree_columns(app.tree_pruefen)
        autosize_tree_columns(app.tree_skip)
    except Exception:
        pass

def _get_sender_address(mail) -> str:
    try:
        sender = mail.Sender
        if sender and sender.AddressEntryUserType == 0:
            ex_user = sender.GetExchangeUser()
            if ex_user:
                return ex_user.PrimarySmtpAddress
    except Exception:
        pass
    try:
        return mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
    except Exception:
        pass
    return str(getattr(mail, "SenderEmailAddress", "") or "")

def send_managers(app, mode: str | None = None) -> None:
    """Controller: Versand für ausgewählte Vorgesetzte auslösen.

    Übernimmt die bisherige Logik aus `App.on_send_managers`, arbeitet aber
    ausschließlich über die `app`-Instanz (UI/State/Tracking).
    """
    sel = app.tree.selection()
    if not sel:
        raise ValueError("Bitte mindestens eine/n Vorgesetzte/n auswählen.")

    rb_year = app.rb_year_var.get()
    ab_year = app.ab_year_var.get()
    # Korrektur: Von services/ aus 2 Ebenen hoch zur Root, dann tracking/versand
    out_root = Path(__file__).parent.parent.parent / "tracking" / "versand"
    out_root.mkdir(parents=True, exist_ok=True)

    # Fortschritt initialisieren
    try:
        total = len(sel)
    except Exception:
        total = 0
    if hasattr(app, "ms_progress"):
        app.ms_progress.configure(mode="determinate", maximum=max(total, 1), value=0)
        app.ms_status.config(text="Starte Versand...", foreground="black")
        app.update_idletasks()

    errors: list[str] = []
    done = 0
    for item_id in sel:
        # Echtes VG-PN aus den Spaltenwerten lesen (iid kann "<PN>_<index>" sein)
        try:
            _mgr_values = app.tree.item(item_id, "values")
            vg_pn = str(_mgr_values[0]) if _mgr_values else ""
        except Exception:
            vg_pn = ""
        pack = app.mgr_index.get(vg_pn)
        if not pack:
            errors.append(f"Kein Paket für VG_PN {vg_pn}")
            done += 1
            if hasattr(app, "ms_progress"):
                app.ms_progress['value'] = done
                app.ms_status.config(text=f"{done}/{total}: Paket fehlt für {vg_pn}")
                app.update_idletasks()
            continue
        mgr = pack["manager"]
        subs = pack["subs"]
        if mgr is None:
            errors.append(f"Vorgesetzte/r mit PN {vg_pn} nicht in EXPORT.xlsx gefunden.")
            done += 1
            if hasattr(app, "ms_progress"):
                app.ms_progress['value'] = done
                app.ms_status.config(text=f"{done}/{total}: Datensatz fehlt für {vg_pn}")
                app.update_idletasks()
            continue

        try:
            if hasattr(app, "ms_progress"):
                mgr_name_lbl = f"{mgr.get('Rufname','')} {mgr.get('Nachname','')}".strip()
                app.ms_status.config(text=f"{done+1}/{total}: Erzeuge & versende an {mgr_name_lbl}...")
                app.update_idletasks()
            build_and_send_for_manager(
                mgr_row=mgr,
                subs_df=subs,
                rb_year=rb_year,
                ab_year=ab_year,
                today=date.today(),
                out_root=out_root,
                managers_index=app.mgr_index,
                include_feedback=True,
                send_mode=mode,
            )

            # Tracking: Logge Versand für jeden Mitarbeiter
            mgr_name = f"{mgr.get('Rufname','')} {mgr.get('Nachname','')}"

            # Erst: Feedback einmal pro Vorgesetzten loggen
            app.tracking.log_feedback_for_manager(vg_pn, mgr_name, rb_year, app.mgr_index)

            # Dann: Dokumente pro Mitarbeiter loggen
            for _, emp in subs.iterrows():
                emp_pn = str(emp.get("ID_NO_ZERO", "")).strip()
                emp_name = f"{emp.get('Rufname','')} {emp.get('Nachname','')}"

                # Bestimme Dokumenttypen basierend auf Mitarbeiter-Status
                # Wichtig: Verwende dieselbe Logik wie determine_docset() aus dispatch_service
                from app.services.dispatch_service import determine_docset
                actual_doc_types = determine_docset(emp, date.today())
                
                # Filtere nur tracking-relevante Typen (Rückblick_Probezeit wird nicht getrackt)
                doc_types: list[str] = []
                for doc_type in actual_doc_types:
                    if doc_type == "Rückblick":
                        doc_types.append("rueckblick")
                    elif doc_type == "Ausblick":
                        doc_types.append("ausblick")
                    elif doc_type == "Rückblick_Probezeit":
                        # Rückblick_Probezeit wird NICHT getrackt (separater Prozess)
                        pass

                # Nur loggen wenn es tracking-relevante Dokumente gibt
                if doc_types:
                    app.tracking.log_versand(
                        mgr_pn=vg_pn,
                        mgr_name=mgr_name,
                        emp_pn=emp_pn,
                        emp_name=emp_name,
                        doc_types=doc_types,
                        rb_year=rb_year,
                        ab_year=ab_year,
                        include_feedback=False,
                    )

        except Exception as e:
            errors.append(f"{mgr.get('Rufname','')} {mgr.get('Nachname','')} ({vg_pn}): {e}")
        finally:
            done += 1
            if hasattr(app, "ms_progress"):
                app.ms_progress['value'] = done
                app.update_idletasks()

    if errors:
        raise RuntimeError("\n".join(errors))
    else:
        pass  # Erfolg im Controller anzeigen

    # Fortschritt abschließen
    if hasattr(app, "ms_progress"):
        app.ms_status.config(text="Bereit.", foreground="gray")
        app.ms_progress['value'] = 0

def send_selected_employees(app, mode: str | None = None) -> None:
    """Controller: Versand für ausgewählte Mitarbeitende eines VG auslösen."""
    sel_mgrs = app.tree_einzel.selection()
    if len(sel_mgrs) != 1:
        raise ValueError("Bitte genau eine/n Vorgesetzte/n auswählen.")
    # Echtes VG-PN aus den Spaltenwerten lesen
    try:
        _mgr_values = app.tree_einzel.item(sel_mgrs[0], "values")
        vg_pn = str(_mgr_values[0]) if _mgr_values else ""
    except Exception:
        vg_pn = ""
    pack = app.mgr_index.get(vg_pn)
    if not pack:
        raise LookupError(f"Kein Paket für VG_PN {vg_pn}")

    subs = pack["subs"].copy()
    sel_subs = app.subs_tree.selection()
    if not sel_subs:
        raise ValueError("Bitte mindestens eine/n Mitarbeitende/n auswählen.")
    # Aus den ausgewählten MA-Zeilen die echte PN aus values[0] lesen
    sel_pns: set[str] = set()
    for iid in sel_subs:
        try:
            vals = app.subs_tree.item(iid, "values")
            pn_val = str(vals[0]) if vals else ""
        except Exception:
            pn_val = ""
        if pn_val:
            sel_pns.add(pn_val)
    if "ID_NO_ZERO" not in subs.columns:
        raise ValueError("Stammdaten enthalten keine Spalte ID_NO_ZERO.")
    subs_filtered = subs[subs["ID_NO_ZERO"].astype(str).isin(sel_pns)]
    if subs_filtered.empty:
        raise LookupError("Keine passenden Mitarbeitenden gefunden.")

    types: list[str] = []
    if app.var_rb.get():
        types.append("Rückblick")
    if app.var_ab.get():
        types.append("Ausblick")
    if app.var_pz.get():
        types.append("Rückblick_Probezeit")
    if not types:
        raise ValueError("Bitte mindestens einen Dokumenttyp auswählen.")

    rb_year = app.rb_year_var_einzel.get()
    ab_year = app.ab_year_var_einzel.get()
    # Korrektur: Von services/ aus 2 Ebenen hoch zur Root, dann tracking/versand
    out_root = Path(__file__).parent.parent.parent / "tracking" / "versand"
    out_root.mkdir(parents=True, exist_ok=True)

    mgr = pack["manager"]
    try:
        if hasattr(app, "es_progress"):
            app.es_progress.configure(mode="indeterminate")
            app.es_progress.start(50)
            app.es_status.config(text="Erzeuge & versende...", foreground="black")
            app.update_idletasks()
        build_and_send_for_manager(
            mgr_row=mgr,
            subs_df=subs_filtered,
            rb_year=rb_year,
            ab_year=ab_year,
            today=date.today(),
            out_root=out_root,
            managers_index=app.mgr_index,
            doc_types_override=types,
            include_feedback=False,
            send_mode=mode,
        )

        mgr_name = f"{mgr.get('Rufname','')} {mgr.get('Nachname','')}"

        # Dokumente pro Mitarbeiter loggen (ohne Feedback)
        for _, emp in subs_filtered.iterrows():
            emp_pn = str(emp.get("ID_NO_ZERO", "")).strip()
            emp_name = f"{emp.get('Rufname','')} {emp.get('Nachname','')}"

            # Filtere nur tracking-relevante Typen (Rückblick_Probezeit wird nicht getrackt)
            doc_types: list[str] = []
            for ui_type in types:
                if ui_type == "Rückblick":
                    doc_types.append("rueckblick")
                elif ui_type == "Ausblick":
                    doc_types.append("ausblick")
                elif ui_type == "Rückblick_Probezeit":
                    # Rückblick_Probezeit wird NICHT getrackt (separater Prozess)
                    pass

            # Nur loggen wenn es tracking-relevante Dokumente gibt
            if doc_types:
                app.tracking.log_versand(
                    mgr_pn=vg_pn,
                    mgr_name=mgr_name,
                    emp_pn=emp_pn,
                    emp_name=emp_name,
                    doc_types=doc_types,
                    rb_year=rb_year,
                    ab_year=ab_year,
                    include_feedback=False,
                )

    except Exception as e:
        raise
    finally:
        if hasattr(app, "es_progress"):
            try:
                app.es_progress.stop()
            except Exception:
                pass
            app.es_status.config(text="Bereit.", foreground="gray")

        # Erfolg im Controller anzeigen


def preview_managers(app) -> None:
    """Vorschau für Massenversand an Vorgesetzte."""
    sel = app.tree.selection()
    if not sel:
        raise ValueError("Bitte mindestens eine/n Vorgesetzte/n auswählen.")
    
    from app.data_loader import load_config
    CFG = load_config()
    rb_year = app.rb_year_var.get()
    ab_year = app.ab_year_var.get()
    subject_tpl = CFG.get("mail", {}).get("subject_template", "MD-Unterlagen Durchlauf {rb_year}/{ab_year}")
    body_tpl = CFG.get("mail", {}).get("body_html_template", "")
    
    previews = []
    
    # Für alle ausgewählten Vorgesetzten Vorschau erstellen
    for item_id in sel:
        try:
            _mgr_values = app.tree.item(item_id, "values")
            vg_pn = str(_mgr_values[0]) if _mgr_values else ""
        except Exception:
            continue
            
        pack = app.mgr_index.get(vg_pn)
        if not pack or pack.get("manager") is None:
            continue
            
        mgr = pack["manager"]
        vg_vorname = str(mgr.get("Rufname", "")).strip()
        vg_nachname = str(mgr.get("Nachname", "")).strip()
        to = str(mgr.get("lange ID/Nummer", "")).strip()
        anrede = f"Hallo {vg_vorname}" if vg_vorname else "Hallo"
        
        try:
            subject = subject_tpl.format(rb_year=rb_year, ab_year=ab_year)
        except Exception:
            subject = subject_tpl or ""
        try:
            body = body_tpl.format(anrede=anrede, rb_year=rb_year, ab_year=ab_year)
        except Exception:
            body = body_tpl or ""
        
        previews.append({
            "to": f"{vg_vorname} {vg_nachname} <{to}>",
            "subject": subject,
            "body": body
        })
    
    if not previews:
        raise ValueError("Keine gültigen Vorgesetzten-Daten gefunden.")
    
    _show_preview_dialog("Vorschau Massenversand", previews)


def preview_selected(app) -> None:
    """Vorschau für Einzelversand an ausgewählte Mitarbeiter."""
    sel_mgrs = app.tree_einzel.selection()
    if not sel_mgrs:
        raise ValueError("Bitte mindestens eine/n Vorgesetzte/n auswählen.")
    
    from app.data_loader import load_config
    CFG = load_config()
    rb_year = app.rb_year_var_einzel.get()
    ab_year = app.ab_year_var_einzel.get()
    subject_tpl = CFG.get("mail_underjaehrig", {}).get("subject_template", "MD-Unterlagen")
    body_tpl = CFG.get("mail_underjaehrig", {}).get("body_html_template", "")
    
    previews = []
    
    # Für alle ausgewählten Vorgesetzten Vorschau erstellen
    for item_id in sel_mgrs:
        try:
            _mgr_values = app.tree_einzel.item(item_id, "values")
            vg_pn = str(_mgr_values[0]) if _mgr_values else ""
        except Exception:
            continue
            
        pack = app.mgr_index.get(vg_pn)
        if not pack or pack.get("manager") is None:
            continue
            
        mgr = pack["manager"]
        vg_vorname = str(mgr.get("Rufname", "")).strip()
        vg_nachname = str(mgr.get("Nachname", "")).strip()
        to = str(mgr.get("lange ID/Nummer", "")).strip()
        anrede = f"Hallo {vg_vorname}" if vg_vorname else "Hallo"
        
        try:
            subject = subject_tpl.format(rb_year=rb_year, ab_year=ab_year)
        except Exception:
            subject = subject_tpl or ""
        try:
            body = body_tpl.format(anrede=anrede, rb_year=rb_year, ab_year=ab_year)
        except Exception:
            body = body_tpl or ""
        
        previews.append({
            "to": f"{vg_vorname} {vg_nachname} <{to}>",
            "subject": subject,
            "body": body
        })
    
    if not previews:
        raise ValueError("Keine gültigen Vorgesetzten-Daten gefunden.")
    
    _show_preview_dialog("Vorschau Einzelversand", previews)


def _show_preview_dialog(title: str, previews: list) -> None:
    """Zeigt Vorschau der E-Mails im Browser als HTML (gleicher Stil wie Erinnerungen)."""
    import tempfile
    import webbrowser
    
    # HTML-Dokument mit allen Vorschauen erstellen
    html_parts = [
        "<html><head><meta charset=\"utf-8\">",
        "<style>",
        "body{font-family:Arial, Helvetica, sans-serif;font-size:14px;line-height:1.5;color:#222;margin:0;padding:0;}",
        ".preview-container{max-width:900px;margin:20px auto;padding:20px;}",
        ".email-preview{background:#fff;border:1px solid #ddd;border-radius:8px;padding:20px;margin-bottom:24px;box-shadow:0 2px 4px rgba(0,0,0,0.1);}",
        ".email-header{background:#f5f5f5;padding:12px;border-radius:4px;margin-bottom:16px;}",
        ".email-to{font-weight:600;color:#0066cc;margin-bottom:4px;}",
        ".email-subject{font-weight:600;margin-bottom:4px;}",
        ".email-body{padding:12px 0;}",
        "h1{color:#333;font-size:24px;margin:0 0 20px 0;border-bottom:2px solid #0066cc;padding-bottom:10px;}",
        "hr{border:none;border-top:2px solid #0066cc;margin:20px 0;}",
        "</style></head><body>",
        "<div class=\"preview-container\">",
        f"<h1>{title}</h1>"
    ]
    
    # Jede Preview als separate Box hinzufügen
    for preview in previews:
        html_parts.append("<div class=\"email-preview\">")
        html_parts.append("<div class=\"email-header\">")
        html_parts.append(f"<div class=\"email-to\">An: {preview['to']}</div>")
        html_parts.append(f"<div class=\"email-subject\">Betreff: {preview['subject']}</div>")
        html_parts.append("</div>")
        html_parts.append(f"<div class=\"email-body\">{preview['body']}</div>")
        html_parts.append("</div>")
    
    html_parts.append("</div></body></html>")
    
    # HTML in temporäre Datei schreiben und im Browser öffnen
    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as f:
        f.write("".join(html_parts))
        temp_path = f.name
    
    webbrowser.open(Path(temp_path).as_uri())


def render_mail_preview(subject_tpl: str, body_tpl: str, anrede: str, rb_year: int, ab_year: int) -> None:
    """Rendert eine HTML-Vorschau der E-Mail (Legacy-Funktion, nicht mehr verwendet)."""
    try:
        subject = (subject_tpl or "").format(rb_year=rb_year, ab_year=ab_year)
    except Exception:
        subject = subject_tpl or ""
    try:
        body = (body_tpl or "").format(anrede=anrede, rb_year=rb_year, ab_year=ab_year)
    except Exception:
        body = body_tpl or ""

    html = (
        "<html><head><meta charset=\"utf-8\">"
        "<style>"
        "body{font-family:Arial, Helvetica, sans-serif;font-size:14px;line-height:1.5;color:#222;margin:16px;}"
        "h3{margin:0 0 8px 0;font-weight:600;}"
        "hr{border:none;border-top:1px solid #ccc;margin:12px 0;}"
        "</style></head><body>"
        f"<h3>Betreff: {subject}</h3><hr/>"
        f"{body}"
        "</body></html>"
    )

    import tempfile, webbrowser
    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as f:
        f.write(html)
        temp_path = f.name
    webbrowser.open(Path(temp_path).as_uri())
