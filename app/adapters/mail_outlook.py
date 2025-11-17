from pathlib import Path
import pythoncom
import win32com.client as win32

from app.data_loader import load_config

CFG = load_config()


def _set_from_account(mail, outlook):
    """Setze das Absenderkonto für Gruppenpostfach (Send On Behalf Of).
    Nutzt choose_account_smtp aus config.yaml, z. B. hr@vd.zh.ch.
    """
    choose_smtp = (CFG.get("mail", {}).get("choose_account_smtp") or "").strip()
    
    if not choose_smtp:
        return  # Kein Konto angegeben -> Standardkonto wird verwendet
    
    try:
        mail.SentOnBehalfOfName = choose_smtp
        print(f"📤 Absenderkonto gesetzt auf: {choose_smtp}")
    except Exception as e:
        print(f"⚠️ Fehler beim Setzen des Absenderkontos: {e}")


def send_mail(
    to: str,
    subject: str,
    html_body: str,
    attachments: list[Path] | None = None,
    cc: str | None = None,
    bcc: str | None = None,
    mode_override: str | None = None,
):
    pythoncom.CoInitialize()
    outlook = win32.Dispatch("Outlook.Application")

    try:
        mail = outlook.CreateItem(0)  # olMailItem
        _set_from_account(mail, outlook)

        mail.To = to
        if cc:
            mail.CC = cc
        cfg_bcc = (CFG.get("mail", {}).get("bcc") or "").strip()
        mail.BCC = bcc or cfg_bcc or ""
        mail.Subject = subject
        mail.HTMLBody = html_body

        for att in attachments or []:
            mail.Attachments.Add(str(att))

        mode = (mode_override or CFG.get("mail", {}).get("send_mode") or "send").lower()

        # Debug: welcher Absender wird tatsächlich verwendet?
        try:
            print(f"📧 'Im Auftrag von': {mail.SentOnBehalfOfName}")
        except Exception:
            print("📧 'Im Auftrag von': (nicht gesetzt - Standardkonto wird verwendet)")

        if mode == "display":
            mail.Display(False)
        else:
            mail.Send()
            print("✅ Mail gesendet.")
    finally:
        pythoncom.CoUninitialize()