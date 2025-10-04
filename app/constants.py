"""
Zentrale Konstanten für das MD-Prozess-Tool.

Autor: VD GS HR
"""

from __future__ import annotations
from enum import Enum


class MDConstants:
    """Sammelt zentrale Konstanten für das MD-Tool."""

    # Stammdaten
    REQUIRED_COLS: list[str] = [
        "ID_NO_ZERO",
        "Rufname",
        "Nachname",
        "OE Bez.",
        "OE Kurzb.",
        "Plans. Bez.",
        "lange ID/Nummer",
        "Dir. Vorgesetzter (PN)",
        "BsGrd",
    ]

    # Rücklauf / Mailverarbeitung
    MAILBOX_NAME: str = "VD-GS HR"
    MAIL_INBOX_NAME: str = "Posteingang"
    MAIL_TARGET_FOLDER_NAME: str = "12 Mitarbeitenden-Dialog"
    MD_KEYWORDS: list[str] = ["rückblick", "rueckblick", "ausblick", "feedback"]
    PROBEZEIT_KEYWORD: str = "probezeit"
    ALLOWED_EXTENSIONS: list[str] = [".docx", ".pdf"]
    SIGNATURE_EXTENSIONS: list[str] = [
        ".jpeg", ".jpg", ".png", ".gif", ".bmp", ".tiff", ".svg", ".ico", ".webp",
    ]

    # GUI
    TREE_MIN_WIDTH: int = 80
    TREE_MAX_WIDTH: int = 420
    TREE_PADDING: int = 24

    # Verarbeitung
    PROC_DEFAULT_BATCH: int = 100
    PROC_MONTH_THRESHOLD: int = 4  # bis inkl. April -> Vorjahr
    PROBEZEIT_MONTHS: tuple[int, ...] = (10, 11, 12, 1)

    # Dashboard
    DASH_STATUS_VALUES: list[str] = [
        "", "ausstehend", "erhalten", "prüfung_nötig", "erübrigt"
    ]

    # Pfade/Ordnernamen (relativ zum Projektwurzelverzeichnis)
    RUECKLAUF_DIR: str = "ruecklauf"
    UNVERARBEITET_DIR: str = "unverarbeitet"
    MANUELL_DIR: str = "manuell"
    VERARBEITET_DIR: str = "verarbeitet"
    LOGS_DIR: str = "logs"

    # UI-Messages
    MSG_SUCCESS: str = "Erfolg"
    MSG_ERROR: str = "Fehler"
    MSG_WARNING: str = "Warnung"
    MSG_INFO: str = "Info"
    MSG_FINISHED: str = "Fertig"
    MSG_HINT: str = "Hinweis"


class DocType(str, Enum):
    """Dokumenttypen im MD-Prozess."""
    RUECKBLICK = "Rückblick"
    AUSBLiCK = "Ausblick"
    FEEDBACK = "Feedback"
    PROBEZEIT = "Probezeit"
    FEEDBACK_VORLAGE_PREFIX = "Vorlage_Feedback_an_"


class ProcStatus(str, Enum):
    """Verarbeitungsstatus für Dateien/Ergebnisse."""
    OK = "ok"
    MANUELL = "manuell"
    PRUEFUNG_NOETIG = "prüfung_nötig"
    ERHALTEN = "erhalten"
    AUSSTEHEND = "ausstehend"
    ERUEBRIGT = "erübrigt"


class DashTag(str, Enum):
    """Tag-Namen für Treeview-Farbcodierung im Dashboard."""
    AUSSTEHEND = "status_ausstehend"
    ERHALTEN = "status_erhalten"
    ERUEBRIGT = "status_eruebrigt"
    PRUEFUNG_NOETIG = "status_pruefung_noetig"
    OK = "status_ok"
    MANUELL = "status_manuell"


