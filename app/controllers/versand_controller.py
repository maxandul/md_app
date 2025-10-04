from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
from tkinter import messagebox

from constants import MDConstants
from dispatch import build_and_send_for_manager
from services.email_service import send_managers, send_selected_employees
from services.sap_data_service import create_vg_ma_relationship


def send_managers(app) -> None:
    """Sendet MD-Dokumente an ausgewählte Vorgesetzte."""
    from services.email_service import send_managers as email_send_managers
    email_send_managers(app)


def send_selected_employees(app) -> None:
    """Sendet MD-Dokumente an ausgewählte Mitarbeiter."""
    from services.email_service import send_selected_employees as email_send_selected
    email_send_selected(app)


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
    email_preview_managers(app)


def preview_selected(app) -> None:
    """Zeigt Vorschau der zu versendenden Dokumente für ausgewählte Mitarbeiter."""
    from services.email_service import preview_selected as email_preview_selected
    email_preview_selected(app)


def render_mail_preview(app) -> None:
    """Rendert eine Vorschau der E-Mail-Inhalte."""
    from services.email_service import render_mail_preview as email_render_preview
    email_render_preview(app)