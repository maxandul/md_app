from __future__ import annotations

from pathlib import Path
import os

import win32com.client
from tkinter import messagebox

from constants import MDConstants, ProcStatus
from services.outlook_service import scan_real, _get_sender_address
from views.ui_utils import autosize_tree_columns


def scan_real(app) -> None:
    """Scannt die Shared Mailbox und verarbeitet eingehende MD-Dokumente."""
    from services.outlook_service import scan_real as outlook_scan_real
    outlook_scan_real(app)