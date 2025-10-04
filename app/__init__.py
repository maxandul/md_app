"""
MD-Prozess-Tool - Mitarbeitenden-Dialog Verwaltung

Eine moderne Python-Anwendung für die Verwaltung des Mitarbeitenden-Dialogs:
- SAP Stammdaten-Validierung
- MD-Dokumenten-Generierung und -Versand
- Rücklauf-Verarbeitung und -Tracking
- Dashboard für Status-Übersicht

Architektur:
- MVC-Pattern mit Controllers, Services und Views
- Service-orientierte Architektur
- Zentrale Konstanten und Konfiguration
- Modulare UI-Komponenten

Autor: VD GS HR
Version: 1.0
"""

from .main import App
from .constants import MDConstants, DocType, ProcStatus, DashTag

__version__ = "1.0.0"
__author__ = "VD GS HR"

__all__ = [
    "App",
    "MDConstants", 
    "DocType",
    "ProcStatus",
    "DashTag"
]
