from __future__ import annotations

from constants import MDConstants
from logging_config import get_logger

logger = get_logger()


def send_managers(app) -> None:
    """Sendet MD-Dokumente an ausgewählte Vorgesetzte."""
    from services.email_service import send_managers as email_send_managers
    try:
        logger.info("Versand: Start send_managers")
        email_send_managers(app)
        from tkinter import messagebox
        messagebox.showinfo(MDConstants.MSG_FINISHED, "Versand abgeschlossen.")
        logger.info("Versand: Ende send_managers (ok)")
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror(MDConstants.MSG_ERROR, str(e))
        logger.error("Versand: Fehler send_managers", extra={"error": str(e)})


def send_selected_employees(app) -> None:
    """Sendet MD-Dokumente an ausgewählte Mitarbeiter."""
    from services.email_service import send_selected_employees as email_send_selected
    try:
        logger.info("Versand: Start send_selected_employees")
        email_send_selected(app)
        from tkinter import messagebox
        messagebox.showinfo(MDConstants.MSG_FINISHED, "Einzelversand abgeschlossen.")
        logger.info("Versand: Ende send_selected_employees (ok)")
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror(MDConstants.MSG_ERROR, str(e))
        logger.error("Versand: Fehler send_selected_employees", extra={"error": str(e)})


def create_vg_ma_relationship(app) -> None:
    """Erstellt neue VG-MA-Beziehung in EXPORT.xlsx."""
    from services.sap_data_service import create_vg_ma_relationship as sap_create_relationship
    sap_create_relationship(app)


def refresh_mgr_table(app) -> None:
    """Aktualisiert die VG-Tabelle im Massenversand-Tab."""
    from services.sap_data_service import refresh_mgr_table as sap_refresh_mgr_table
    sap_refresh_mgr_table(app)


def refresh_mgr_table_einzel(app) -> None:
    """Aktualisiert die VG-Tabelle im Einzelversand-Tab."""
    from services.sap_data_service import refresh_mgr_table_einzel as sap_refresh_mgr_table_einzel
    sap_refresh_mgr_table_einzel(app)


def refresh_vg_list(app) -> None:
    """Aktualisiert die VG-Liste im VG-MA-Tab."""
    from services.sap_data_service import refresh_vg_list as sap_refresh_vg_list
    sap_refresh_vg_list(app)


def refresh_ma_list(app) -> None:
    """Aktualisiert die MA-Liste im VG-MA-Tab."""
    from services.sap_data_service import refresh_ma_list as sap_refresh_ma_list
    sap_refresh_ma_list(app)


def update_selection_status(app) -> None:
    """Aktualisiert den Auswahlinfo-Text im VG-MA-Tab."""
    from services.sap_data_service import update_selection_status as sap_update_selection_status
    sap_update_selection_status(app)


def preview_managers(app) -> None:
    """Zeigt Vorschau der zu versendenden Dokumente für ausgewählte Vorgesetzte."""
    from services.email_service import preview_managers as email_preview_managers
    try:
        logger.info("Versand: Start preview_managers")
        email_preview_managers(app)
        logger.info("Versand: Ende preview_managers (ok)")
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror(MDConstants.MSG_ERROR, str(e))
        logger.error("Versand: Fehler preview_managers", extra={"error": str(e)})


def preview_selected(app) -> None:
    """Zeigt Vorschau der zu versendenden Dokumente für ausgewählte Mitarbeiter."""
    from services.email_service import preview_selected as email_preview_selected
    try:
        logger.info("Versand: Start preview_selected")
        email_preview_selected(app)
        logger.info("Versand: Ende preview_selected (ok)")
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror(MDConstants.MSG_ERROR, str(e))
        logger.error("Versand: Fehler preview_selected", extra={"error": str(e)})


def render_mail_preview(app) -> None:
    """Rendert eine Vorschau der E-Mail-Inhalte."""
    from services.email_service import render_mail_preview as email_render_preview
    try:
        logger.info("Versand: Start render_mail_preview")
        email_render_preview(app)
        logger.info("Versand: Ende render_mail_preview (ok)")
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror(MDConstants.MSG_ERROR, str(e))
        logger.error("Versand: Fehler render_mail_preview", extra={"error": str(e)})