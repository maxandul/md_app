from __future__ import annotations

import tkinter as tk
from tkinter import messagebox

from constants import MDConstants


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
    try:
        dashboard_export(app)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror(MDConstants.MSG_ERROR, f"Export fehlgeschlagen: {e}")


def manual_adjustment(app) -> None:
    """Öffnet einen Dialog zur manuellen Anpassung eines ausgewählten Dashboard-Eintrags."""
    from services.dashboard_service import manual_adjustment as dashboard_manual
    try:
        dashboard_manual(app)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror(MDConstants.MSG_ERROR, f"Manuelle Anpassung fehlgeschlagen: {e}")