from __future__ import annotations

import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

from constants import MDConstants, ProcStatus, DashTag
from services.dashboard_service import refresh_dashboard, export_dashboard, manual_adjustment


def refresh_dashboard(app) -> None:
    """Lädt Dashboard-Daten gemäß aktueller Filter und befüllt den Treeview."""
    try:
        from services.dashboard_service import refresh_dashboard as dashboard_refresh
        dashboard_refresh(app)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror(MDConstants.MSG_ERROR, f"Dashboard konnte nicht aktualisiert werden: {e}")


def export_dashboard(app) -> None:
    """Exportiert die aktuell gefilterten Dashboard-Daten als CSV."""
    from services.dashboard_service import export_dashboard as dashboard_export
    dashboard_export(app)


def manual_adjustment(app) -> None:
    """Öffnet einen Dialog zur manuellen Anpassung eines ausgewählten Dashboard-Eintrags."""
    from services.dashboard_service import manual_adjustment as dashboard_manual
    dashboard_manual(app)