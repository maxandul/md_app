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
    """Zeigt einen Dialog mit Vorschau der Erinnerungsmails."""
    dialog = tk.Toplevel(app)
    dialog.title("Vorschau Erinnerungsmails")
    dialog.geometry("800x600")
    dialog.transient(app)
    dialog.grab_set()

    # Notebook für mehrere Vorschauen
    notebook = tk.ttk.Notebook(dialog)
    notebook.pack(fill="both", expand=True, padx=8, pady=8)

    for i, preview in enumerate(previews):
        frame = tk.ttk.Frame(notebook)
        notebook.add(frame, text=f"An: {preview['to']}")

        # Betreff
        tk.ttk.Label(frame, text="Betreff:", font=("Arial", 10, "bold")).pack(anchor="w", padx=8, pady=(8, 4))
        subject_text = tk.Text(frame, height=2, wrap="word")
        subject_text.pack(fill="x", padx=8, pady=(0, 8))
        subject_text.insert("1.0", preview["subject"])
        subject_text.config(state="disabled")

        # Inhalt
        tk.ttk.Label(frame, text="Inhalt:", font=("Arial", 10, "bold")).pack(anchor="w", padx=8, pady=(8, 4))
        content_text = tk.Text(frame, wrap="word")
        content_text.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        
        # HTML zu Text konvertieren für bessere Lesbarkeit
        html_content = preview["body"]
        # Einfache HTML-Tag-Entfernung
        import re
        text_content = re.sub(r'<[^>]+>', '', html_content)
        text_content = text_content.replace('&nbsp;', ' ').replace('&lt;', '<').replace('&gt;', '>')
        
        content_text.insert("1.0", text_content)
        content_text.config(state="disabled")

    # Schließen-Button
    button_frame = tk.ttk.Frame(dialog)
    button_frame.pack(fill="x", padx=8, pady=(0, 8))
    tk.ttk.Button(button_frame, text="Schließen", command=dialog.destroy).pack(side="right")
