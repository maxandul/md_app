"""
Outlook-Service für das MD-Prozess-Tool.

Dieser Service kapselt alle Funktionen zur Outlook-Integration,
einschließlich E-Mail-Scanning und Anhang-Verarbeitung.
"""

from __future__ import annotations
from pathlib import Path
import os
import win32com.client
from tkinter import messagebox

from constants import MDConstants, ProcStatus


def scan_real(app) -> None:
    """Scannt die Shared Mailbox und verarbeitet eingehende MD-Dokumente."""
    for t in [app.tree_ok, app.tree_pruefen, app.tree_skip]:
        for i in t.get_children():
            t.delete(i)

    if hasattr(app, "inbox_status"):
        app.inbox_status.config(text="Scan läuft...", foreground="black")
        app.update_idletasks()

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    mailbox = outlook.Folders[MDConstants.MAILBOX_NAME]
    inbox = mailbox.Folders[MDConstants.MAIL_INBOX_NAME]
    target_folder = inbox.Folders[MDConstants.MAIL_TARGET_FOLDER_NAME]

    md_keywords = MDConstants.MD_KEYWORDS
    kw_probezeit = MDConstants.PROBEZEIT_KEYWORD
    allowed_exts = MDConstants.ALLOWED_EXTENSIONS
    signature_exts = MDConstants.SIGNATURE_EXTENSIONS

    base_path = Path(app.inbox_target_var.get())
    base_path.mkdir(parents=True, exist_ok=True)

    found = copied = moved = to_check = skipped = 0

    for mail in inbox.Items:
        found += 1
        sender = _get_sender_address(mail)
        subject = str(mail.Subject or "")

        files = []
        for att in mail.Attachments:
            try:
                fname = str(att.FileName or "").strip()
            except Exception:
                continue
            if not fname or "." not in fname:
                continue
            files.append((fname, att))

        if not files:
            app.tree_skip.insert("", "end", values=[sender, subject, "Keine Anhänge"])
            skipped += 1
            continue

        md_files = [f for f, _ in files if any(k in f.lower() for k in md_keywords) and os.path.splitext(f)[1].lower() in allowed_exts]
        probezeit_files = [f for f, _ in files if kw_probezeit in f.lower() and os.path.splitext(f)[1].lower() in allowed_exts]
        signature_files = [f for f, _ in files if os.path.splitext(f)[1].lower() in signature_exts]
        other_files = [f for f, _ in files if f not in md_files and f not in probezeit_files and f not in signature_files]

        if md_files and not probezeit_files and not other_files:
            for fname, att in files:
                if fname in md_files:
                    save_path = base_path / fname
                    att.SaveAsFile(str(save_path))
                    app.tree_ok.insert("", "end", values=[fname, str(base_path), sender, subject])
                    copied += 1
            mail.Move(target_folder)
            moved += 1

        elif probezeit_files:
            grund = "Probezeit"
            app.tree_pruefen.insert("", "end", values=[grund, ", ".join(probezeit_files) or "Keine", sender, subject, "Keine"])
            to_check += 1

        elif md_files and other_files:
            copied_names = []
            for fname, att in files:
                if fname in md_files:
                    save_path = base_path / fname
                    att.SaveAsFile(str(save_path))
                    copied += 1
                    copied_names.append(fname)
            grund = "Fremde Anhänge"
            app.tree_pruefen.insert("", "end", values=[grund, ", ".join(other_files) or "Keine", sender, subject, ", ".join(copied_names) or "Keine"])
            to_check += 1

        elif probezeit_files and other_files:
            grund = "Probezeit + Fremde Anhänge"
            zu_pruefen = probezeit_files + other_files
            app.tree_pruefen.insert("", "end", values=[grund, ", ".join(zu_pruefen) or "Keine", sender, subject, "Keine"])
            to_check += 1

        else:
            app.tree_skip.insert("", "end", values=[sender, subject, "Keine MD-Anhänge"])
            skipped += 1

        if hasattr(app, "inbox_status"):
            app.inbox_status.config(text=f"Scan läuft... geprüft: {found} • kopiert: {copied} • verschoben: {moved} • prüfen: {to_check} • übersprungen: {skipped}")
            app.update_idletasks()

    app.ruecklauf_status.config(
        text=f"Scan abgeschlossen: {found} Mails • {copied} Anhänge kopiert • {moved} verschoben • {to_check} prüfen • {skipped} übersprungen",
        foreground="black",
    )


def _get_sender_address(mail) -> str:
    """Extrahiert die E-Mail-Adresse des Absenders aus einem Outlook-Mail-Objekt."""
    try:
        return str(mail.SenderEmailAddress or mail.SenderName or "Unbekannt")
    except Exception:
        return "Unbekannt"


def process_mail_attachments(mail, base_path: Path) -> tuple[int, int, int, int, int]:
    """
    Verarbeitet Anhänge einer E-Mail.
    
    Returns:
        Tuple mit (found, copied, moved, to_check, skipped)
    """
    # Optional: Kann später implementiert werden
    pass
