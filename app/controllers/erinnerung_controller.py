from __future__ import annotations

import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta

from app.constants import MDConstants
from app.services.erinnerung_service import ErinnerungService


def send_reminders(app) -> None:
    """Versendet Erinnerungsmails für ausgewählte Dashboard-Einträge."""
    selection = app.tree_dashboard.selection()
    if not selection:
        messagebox.showwarning(MDConstants.MSG_WARNING, "Bitte wählen Sie mindestens einen Eintrag aus.")
        return

    # Bestätigungsdialog
    count = len(selection)
    if not messagebox.askyesno(
        "Erinnerung versenden", 
        f"Möchten Sie wirklich Erinnerungen für {count} ausgewählte Einträge versenden?"
    ):
        return

    try:
        service = ErinnerungService(app)
        result = service.send_reminders(selection, mode="send")
        
        if result["success"]:
            messagebox.showinfo(
                MDConstants.MSG_SUCCESS, 
                f"Erinnerungen erfolgreich versendet!\n"
                f"E-Mails gesendet: {result['emails_sent']}\n"
                f"Fehler: {result['errors']}"
            )
            # Dashboard aktualisieren um neue Erinnerungsdaten zu zeigen
            from app.controllers.dashboard_controller import refresh_dashboard
            refresh_dashboard(app)
        else:
            messagebox.showerror(MDConstants.MSG_ERROR, f"Fehler beim Versenden der Erinnerungen: {result['error']}")
            
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Unerwarteter Fehler: {e}")


def save_reminders_as_draft(app) -> None:
    """Speichert Erinnerungsmails als Entwürfe für ausgewählte Dashboard-Einträge."""
    selection = app.tree_dashboard.selection()
    if not selection:
        messagebox.showwarning(MDConstants.MSG_WARNING, "Bitte wählen Sie mindestens einen Eintrag aus.")
        return

    # Bestätigungsdialog
    count = len(selection)
    if not messagebox.askyesno(
        "Erinnerung als Entwurf speichern", 
        f"Möchten Sie wirklich Erinnerungs-Entwürfe für {count} ausgewählte Einträge erstellen?"
    ):
        return

    try:
        service = ErinnerungService(app)
        result = service.send_reminders(selection, mode="display")
        
        if result["success"]:
            messagebox.showinfo(
                MDConstants.MSG_SUCCESS, 
                f"Erinnerungs-Entwürfe erfolgreich erstellt!\n"
                f"Entwürfe erstellt: {result['emails_sent']}\n"
                f"Fehler: {result['errors']}"
            )
        else:
            messagebox.showerror(MDConstants.MSG_ERROR, f"Fehler beim Erstellen der Entwürfe: {result['error']}")
            
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Unerwarteter Fehler: {e}")


def preview_reminders(app) -> None:
    """Zeigt Vorschau der Erinnerungsmails für ausgewählte Dashboard-Einträge."""
    selection = app.tree_dashboard.selection()
    if not selection:
        messagebox.showwarning(MDConstants.MSG_WARNING, "Bitte wählen Sie mindestens einen Eintrag aus.")
        return

    try:
        service = ErinnerungService(app)
        result = service.preview_reminders(selection)
        
        if result["success"]:
            _show_preview_dialog(app, result["previews"])
        else:
            messagebox.showerror(MDConstants.MSG_ERROR, f"Fehler beim Generieren der Vorschau: {result['error']}")
            
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Unerwarteter Fehler: {e}")


def _show_preview_dialog(app, previews: list) -> None:
    """Zeigt Vorschau der Erinnerungsmails im Browser als HTML."""
    import tempfile
    import webbrowser
    from pathlib import Path
    
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
        "<h1>Vorschau Erinnerungsmails</h1>"
    ]
    
    # Jede Preview als separate Box hinzufügen
    for i, preview in enumerate(previews):
        html_parts.append("<div class=\"email-preview\">")
        html_parts.append("<div class=\"email-header\">")
        html_parts.append(f"<div class=\"email-to\">An: {preview['to']}</div>")
        html_parts.append(f"<div class=\"email-subject\">Betreff: {preview['subject']}</div>")
        html_parts.append("</div>")
        html_parts.append("<div class=\"email-body\">")
        html_parts.append(preview["body"])
        html_parts.append("</div>")
        html_parts.append("</div>")
        
        # Trennlinie zwischen Emails (außer nach der letzten)
        if i < len(previews) - 1:
            html_parts.append("<hr/>")
    
    html_parts.append("</div></body></html>")
    
    # HTML-Datei temporär speichern und im Browser öffnen
    html_content = "\n".join(html_parts)
    
    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as f:
        f.write(html_content)
        temp_path = f.name
    
    webbrowser.open(Path(temp_path).as_uri())
