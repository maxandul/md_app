# app/mail_send.py
from pathlib import Path
import pythoncom
import win32com.client as win32
from data_loader import load_config

CFG = load_config()

def _set_from_account(mail, outlook):
    """Optional: Sendeadresse setzen.
    - from_address -> SentOnBehalfOfName
    - choose_account_smtp -> wähle passendes Konto aus Outlook.Accounts
    """
    from_addr = (CFG.get("mail", {}).get("from_address") or "").strip()
    choose_smtp = (CFG.get("mail", {}).get("choose_account_smtp") or "").strip()

    if from_addr:
        # Senden im Auftrag von
        mail.SentOnBehalfOfName = from_addr

    if choose_smtp:
        try:
            session = outlook.Session
            accounts = session.Accounts
            for i in range(1, accounts.Count + 1):
                acc = accounts.Item(i)
                if getattr(acc, "SmtpAddress", "") and acc.SmtpAddress.lower() == choose_smtp.lower():
                    mail.SendUsingAccount = acc
                    break
        except Exception:
            pass  # Fallback: Outlook default

def send_mail(to: str, subject: str, html_body: str, attachments: list[Path] = None, cc: str | None = None, bcc: str | None = None, mode_override: str | None = None):
    pythoncom.CoInitialize()
    outlook = win32.Dispatch("Outlook.Application")
    try:
        mail = outlook.CreateItem(0)  # olMailItem
        _set_from_account(mail, outlook)

        mail.To = to
        if cc: mail.CC = cc
        cfg_bcc = (CFG.get("mail", {}).get("bcc") or "").strip()
        mail.BCC = bcc or cfg_bcc or ""
        mail.Subject = subject
        mail.HTMLBody = html_body

        # Reply-To (nur per Header-Trick, optional)
        reply_to = (CFG.get("mail", {}).get("reply_to") or "").strip()
        if reply_to:
            # PR_REPLY_RECIPIENT_NAMES etc. sind über MAPI komplizierter; einfacher: im Body erwähnen
            pass

        for att in attachments or []:
            mail.Attachments.Add(str(att))

        # Modus: optionaler Override pro Aufruf, sonst aus config
        mode = (mode_override or CFG.get("mail", {}).get("send_mode") or "send").lower()
        if mode == "display":
            mail.Display(False)
        else:
            mail.Send()
    finally:
        pythoncom.CoUninitialize()
