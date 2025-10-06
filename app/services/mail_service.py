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

from constants import MDConstants, ProcStatus
from dispatch import build_and_send_for_manager
from data_loader import load_employees, build_manager_index, load_config
from views.ui_utils import autosize_tree_columns


def _get_sender_address(mail) -> str:
    """Extrahiert die primäre SMTP-Adresse des Absenders.

    Versucht mehrere Methoden in sinnvoller Reihenfolge, um statt LegacyDN
    (z. B. "/O=EXCHANGELABS/…") die echte SMTP-Adresse zu erhalten.
    """
    # 1) Über Sender -> ExchangeUser -> PrimarySmtpAddress
    try:
        sender = getattr(mail, "Sender", None)
        if sender and getattr(sender, "AddressEntryUserType", None) == 0:
            ex_user = sender.GetExchangeUser()
            if ex_user:
                addr = getattr(ex_user, "PrimarySmtpAddress", "")
                if addr:
                    return str(addr)
    except Exception:
        pass

    # 2) Über MAPI PropertyAccessor: PR_SMTP_ADDRESS (0x39FE001E)
    try:
        prop = mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
        if prop:
            return str(prop)
    except Exception:
        pass

    # 3) Fallbacks
    try:
        if getattr(mail, "SenderEmailAddress", None):
            return str(mail.SenderEmailAddress)
        if getattr(mail, "SenderName", None):
            return str(mail.SenderName)
    except Exception:
        pass
    return "Unbekannt"


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


def send_managers(app) -> None:
    """Sendet MD-Dokumente an ausgewählte Vorgesetzte."""
    selected_items = app.tree.selection()
    if not selected_items:
        messagebox.showwarning(MDConstants.MSG_WARNING, "Bitte wählen Sie mindestens einen Vorgesetzten aus.")
        return
    
    try:
        rb_year = app.rb_year_var.get()
        ab_year = app.ab_year_var.get()
        today = date.today()
        
        # Lade Konfiguration
        CFG = load_config()
        out_root = Path(CFG["paths"]["output_dir"])
        
        sent_count = 0
        for item_id in selected_items:
            vg_pn = item_id
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
                    send_mode=None
                )
                sent_count += 1
            except Exception as e:
                messagebox.showerror(MDConstants.MSG_ERROR, f"Fehler beim Versand an {mgr.get('Nachname', '')} {mgr.get('Rufname', '')}: {e}")
        
        messagebox.showinfo(MDConstants.MSG_SUCCESS, f"Versand abgeschlossen: {sent_count} Vorgesetzte erhalten MD-Dokumente.")
        
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Fehler beim Versand: {e}")


def send_selected_employees(app) -> None:
    """Sendet MD-Dokumente an ausgewählte Mitarbeiter."""
    selected_items = app.tree_einzel.selection()
    if not selected_items:
        messagebox.showwarning(MDConstants.MSG_WARNING, "Bitte wählen Sie mindestens einen Mitarbeiter aus.")
        return
    
    try:
        rb_year = app.rb_year_var_einzel.get()
        ab_year = app.ab_year_var_einzel.get()
        today = date.today()
        
        # Lade Konfiguration
        CFG = load_config()
        out_root = Path(CFG["paths"]["output_dir"])
        
        sent_count = 0
        for item_id in selected_items:
            vg_pn = item_id
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
                    send_mode=None
                )
                sent_count += 1
            except Exception as e:
                messagebox.showerror(MDConstants.MSG_ERROR, f"Fehler beim Versand an {mgr.get('Nachname', '')} {mgr.get('Rufname', '')}: {e}")
        
        messagebox.showinfo(MDConstants.MSG_SUCCESS, f"Versand abgeschlossen: {sent_count} Vorgesetzte erhalten MD-Dokumente.")
        
    except Exception as e:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Fehler beim Versand: {e}")


def preview_managers(app) -> None:
    """Zeigt Vorschau der zu versendenden Dokumente für ausgewählte Vorgesetzte."""
    selected_items = app.tree.selection()
    if not selected_items:
        messagebox.showwarning(MDConstants.MSG_WARNING, "Bitte wählen Sie mindestens einen Vorgesetzten aus.")
        return
    
    preview_text = "Vorschau der zu versendenden Dokumente:\n\n"
    
    for item_id in selected_items:
        vg_pn = item_id
        if vg_pn not in app.mgr_index:
            continue
            
        pack = app.mgr_index[vg_pn]
        mgr = pack["manager"]
        subs = pack["subs"]
        
        if mgr is None or subs.empty:
            continue
            
        preview_text += f"Vorgesetzter: {mgr.get('Nachname', '')} {mgr.get('Rufname', '')} ({vg_pn})\n"
        preview_text += f"E-Mail: {mgr.get('lange ID/Nummer', '')}\n"
        preview_text += f"Mitarbeiter: {len(subs)} Personen\n\n"
        
        for _, sub in subs.iterrows():
            preview_text += f"  - {sub.get('Nachname', '')} {sub.get('Rufname', '')} ({sub.get('ID_NO_ZERO', '')})\n"
        preview_text += "\n"
    
    # Zeige Vorschau in einem separaten Fenster
    preview_window = tk.Toplevel(app)
    preview_window.title("Vorschau - MD-Versand")
    preview_window.geometry("600x400")
    
    text_widget = tk.Text(preview_window, wrap=tk.WORD)
    scrollbar = ttk.Scrollbar(preview_window, orient="vertical", command=text_widget.yview)
    text_widget.configure(yscrollcommand=scrollbar.set)
    
    text_widget.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    text_widget.insert("1.0", preview_text)
    text_widget.config(state="disabled")


def preview_selected(app) -> None:
    """Zeigt Vorschau der zu versendenden Dokumente für ausgewählte Mitarbeiter."""
    selected_items = app.tree_einzel.selection()
    if not selected_items:
        messagebox.showwarning(MDConstants.MSG_WARNING, "Bitte wählen Sie mindestens einen Mitarbeiter aus.")
        return
    
    preview_text = "Vorschau der zu versendenden Dokumente:\n\n"
    
    for item_id in selected_items:
        vg_pn = item_id
        if vg_pn not in app.mgr_index:
            continue
            
        pack = app.mgr_index[vg_pn]
        mgr = pack["manager"]
        subs = pack["subs"]
        
        if mgr is None or subs.empty:
            continue
            
        preview_text += f"Vorgesetzter: {mgr.get('Nachname', '')} {mgr.get('Rufname', '')} ({vg_pn})\n"
        preview_text += f"E-Mail: {mgr.get('lange ID/Nummer', '')}\n"
        preview_text += f"Mitarbeiter: {len(subs)} Personen\n\n"
        
        for _, sub in subs.iterrows():
            preview_text += f"  - {sub.get('Nachname', '')} {sub.get('Rufname', '')} ({sub.get('ID_NO_ZERO', '')})\n"
        preview_text += "\n"
    
    # Zeige Vorschau in einem separaten Fenster
    preview_window = tk.Toplevel(app)
    preview_window.title("Vorschau - MD-Versand")
    preview_window.geometry("600x400")
    
    text_widget = tk.Text(preview_window, wrap=tk.WORD)
    scrollbar = ttk.Scrollbar(preview_window, orient="vertical", command=text_widget.yview)
    text_widget.configure(yscrollcommand=scrollbar.set)
    
    text_widget.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    text_widget.insert("1.0", preview_text)
    text_widget.config(state="disabled")


def render_mail_preview(app) -> None:
    """Rendert eine Vorschau der E-Mail-Inhalte."""
    # Diese Funktion kann erweitert werden, um eine detaillierte E-Mail-Vorschau zu zeigen
    messagebox.showinfo(MDConstants.MSG_INFO, "E-Mail-Vorschau wird geladen...")

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
        messagebox.showwarning(MDConstants.MSG_HINT, "Bitte mindestens eine/n Vorgesetzte/n auswählen.")
        return

    rb_year = app.rb_year_var.get()
    ab_year = app.ab_year_var.get()
    out_root = Path(__file__).parent.parent / "tracking" / "versand"
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
    for vg_pn in sel:
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
                doc_types: list[str] = []

                # Rückblick für alle
                doc_types.append("rueckblick")

                # Ausblick nur wenn nicht austretend
                if pd.isna(emp.get("Austritt")) or emp.get("Austritt") == "":
                    doc_types.append("ausblick")

                # Probezeit-Rückblick wenn Probezeit Ende zwischen Okt-Jan
                # HINWEIS: rueckblick_probezeit wird nicht im Tracking erfasst (separater Prozess)
                if not pd.isna(emp.get("Ende Probezeit")):
                    probezeit_ende = emp.get("Ende Probezeit")
                    if isinstance(probezeit_ende, pd.Timestamp):
                        month = probezeit_ende.month
                        if month in MDConstants.PROBEZEIT_MONTHS:
                            # doc_types.append("rueckblick_probezeit")  # Auskommentiert: separater Prozess
                            pass

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
        messagebox.showerror(MDConstants.MSG_ERROR, "\n".join(errors))
    else:
        messagebox.showinfo(MDConstants.MSG_FINISHED, "Versand ausgefuehrt (siehe Outbox/gesendete Elemente).")

    # Fortschritt abschließen
    if hasattr(app, "ms_progress"):
        app.ms_status.config(text="Bereit.", foreground="gray")
        app.ms_progress['value'] = 0

def send_selected_employees(app, mode: str | None = None) -> None:
    """Controller: Versand für ausgewählte Mitarbeitende eines VG auslösen."""
    sel_mgrs = app.tree_einzel.selection()
    if len(sel_mgrs) != 1:
        messagebox.showwarning(MDConstants.MSG_HINT, "Bitte genau eine/n Vorgesetzte/n auswählen.")
        return
    vg_pn = sel_mgrs[0]
    pack = app.mgr_index.get(vg_pn)
    if not pack:
        messagebox.showerror(MDConstants.MSG_ERROR, f"Kein Paket für VG_PN {vg_pn}")
        return

    subs = pack["subs"].copy()
    sel_subs = app.subs_tree.selection()
    if not sel_subs:
        messagebox.showwarning(MDConstants.MSG_HINT, "Bitte mindestens eine/n Mitarbeitende/n auswählen.")
        return

    sel_pns = set(sel_subs)
    if "ID_NO_ZERO" not in subs.columns:
        messagebox.showerror(MDConstants.MSG_ERROR, "Stammdaten enthalten keine Spalte ID_NO_ZERO.")
        return
    subs_filtered = subs[subs["ID_NO_ZERO"].astype(str).isin(sel_pns)]
    if subs_filtered.empty:
        messagebox.showerror(MDConstants.MSG_ERROR, "Keine passenden Mitarbeitenden gefunden.")
        return

    types: list[str] = []
    if app.var_rb.get():
        types.append("Rückblick")
    if app.var_ab.get():
        types.append("Ausblick")
    if app.var_pz.get():
        types.append("Rückblick_Probezeit")
    if not types:
        messagebox.showwarning(MDConstants.MSG_HINT, "Bitte mindestens einen Dokumenttyp auswählen.")
        return

    rb_year = app.rb_year_var_einzel.get()
    ab_year = app.ab_year_var_einzel.get()
    out_root = Path(__file__).parent.parent / "tracking" / "versand"
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

            doc_types: list[str] = []
            for ui_type in types:
                if ui_type == "Rückblick":
                    doc_types.append("rueckblick")
                elif ui_type == "Ausblick":
                    doc_types.append("ausblick")
                elif ui_type == "Rückblick_Probezeit":
                    # doc_types.append("rueckblick_probezeit")  # Auskommentiert: separater Prozess
                    pass

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
        messagebox.showerror(MDConstants.MSG_ERROR, str(e))
        return
    finally:
        if hasattr(app, "es_progress"):
            try:
                app.es_progress.stop()
            except Exception:
                pass
            app.es_status.config(text="Bereit.", foreground="gray")

        messagebox.showinfo(MDConstants.MSG_FINISHED, "Unterlagen erzeugt und E-Mail vorbereitet/gesendet.")


def preview_managers(app) -> None:
    """Vorschau für Massenversand an Vorgesetzte."""
    sel = app.tree.selection()
    if not sel:
        messagebox.showwarning(MDConstants.MSG_HINT, "Bitte mindestens eine/n Vorgesetzte/n auswählen.")
        return
    vg_pn = sel[0]
    pack = app.mgr_index.get(vg_pn)
    if not pack or pack.get("manager") is None:
        messagebox.showerror("Fehler", "Kein gültiger Vorgesetzten-Datensatz gefunden.")
        return

    mgr = pack["manager"]
    vg_vorname = str(mgr.get("Rufname", "")).strip()
    anrede = f"Hallo {vg_vorname}" if vg_vorname else "Hallo"

    rb_year = app.rb_year_var.get()
    ab_year = app.ab_year_var.get()

    from data_loader import load_config
    CFG = load_config()
    subject_tpl = CFG.get("mail", {}).get("subject_template", "MD-Unterlagen Durchlauf {rb_year}/{ab_year}")
    body_tpl = CFG.get("mail", {}).get("body_html_template", "")

    render_mail_preview(subject_tpl, body_tpl, anrede, rb_year, ab_year)


def preview_selected(app) -> None:
    """Vorschau für Einzelversand an ausgewählte Mitarbeiter."""
    sel_mgrs = app.tree_einzel.selection()
    if len(sel_mgrs) != 1:
        messagebox.showwarning(MDConstants.MSG_HINT, "Bitte genau eine/n Vorgesetzte/n auswählen.")
        return
    vg_pn = sel_mgrs[0]
    pack = app.mgr_index.get(vg_pn)
    if not pack or pack.get("manager") is None:
        messagebox.showerror("Fehler", "Kein gültiger Vorgesetzten-Datensatz gefunden.")
        return

    mgr = pack["manager"]
    vg_vorname = str(mgr.get("Rufname", "")).strip()
    anrede = f"Hallo {vg_vorname}" if vg_vorname else "Hallo"

    rb_year = app.rb_year_var_einzel.get()
    ab_year = app.ab_year_var_einzel.get()

    from data_loader import load_config
    CFG = load_config()
    subject_tpl = CFG.get("mail_underjaehrig", {}).get("subject_template", "MD-Unterlagen")
    body_tpl = CFG.get("mail_underjaehrig", {}).get("body_html_template", "")

    render_mail_preview(subject_tpl, body_tpl, anrede, rb_year, ab_year)


def render_mail_preview(subject_tpl: str, body_tpl: str, anrede: str, rb_year: int, ab_year: int) -> None:
    """Rendert eine HTML-Vorschau der E-Mail."""
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
