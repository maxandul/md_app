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

# Vorsicht: Import von App triggert viele weitere Importe. Für Paket-Importe
# vermeiden wir Seiteneffekte und exportieren nur Konstanten. Die App sollte
# über `python -m app.main` oder `from app.main import App` geladen werden.
from .constants import MDConstants, DocType, ProcStatus, DashTag

__version__ = "1.0.0"
__author__ = "VD GS HR"

__all__ = [
    "MDConstants", 
    "DocType",
    "ProcStatus",
    "DashTag"
]
