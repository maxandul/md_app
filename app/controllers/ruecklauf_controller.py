from __future__ import annotations

from constants import MDConstants


def scan_real(app) -> None:
    """Scannt die Shared Mailbox und verarbeitet eingehende MD-Dokumente."""
    from services.outlook_service import scan_real as outlook_scan_real
    try:
        outlook_scan_real(app)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror(MDConstants.MSG_ERROR, str(e))