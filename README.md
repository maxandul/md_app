# MD-Prozess-Tool

Ein Python-basiertes Tool zur Verwaltung des Mitarbeitenden-Dialog (MD) Prozesses.

## 📋 Übersicht

Das MD-Prozess-Tool automatisiert die Verwaltung des jährlichen Mitarbeitenden-Dialogs:

- **SAP Stammdaten-Validierung**: Prüfung und Validierung der Datei EXPORT.xlsx (SAP ad-hoc Query "VD_MD")
- **Dokumenten-Generierung**: Automatische Erstellung von MD-Dokumenten (Rückblick/Ausblick/Probezeit/Feedback Vorlage)
- **E-Mail-Versand**: Versendung der generierten Dokumente an Vorgesetzte
- **Rücklauf-Verarbeitung**: Automatische Verarbeitung eingehender MD-Dokumente
- **Status-Tracking**: Dashboard für Übersicht über den MD-Prozess

## 🏗️ Projektstruktur

```
md_app/
├── app/                         # Hauptanwendung
│   ├── main.py                  # GUI-Hauptanwendung (Tkinter)
│   ├── config.yaml              # Konfigurationsdatei
│   ├── logging_config.py        # Zentrales Logging
│   ├── exceptions.py            # Eigene Exception-Typen
│   ├── utils.py                 # Hilfsfunktionen (Dateinamen, Dialoge)
│   ├── data_loader.py           # Konfiguration & SAP-Daten laden, Validierung
│   ├── adapters/                # Technische Adapter (COM/Datei)
│   │   ├── mail_outlook.py      # Outlook E-Mail Versand
│   │   ├── word_outlook.py      # Word-Template Befüllung
│   │   └── docx_reader.py       # Lesen von DOCX-Content-Controls
│   ├── services/                # Geschäftslogik
│   │   ├── document_service.py  # DOCX/PDF-Verarbeitung, Exporte orchestrieren
│   │   ├── email_service.py     # Versand-Flows (nutzt dispatch_service)
│   │   ├── outlook_service.py   # Outlook-Scan (Eingang)
│   │   ├── dispatch_service.py  # Dokumente generieren & Mail vorbereiten
│   │   ├── export_service.py    # SAP-/DS-Exporte
│   │   ├── file_service.py      # Verschiebe-Logik (OK/manuell)
│   │   ├── sap_data_service.py  # Stammdaten-Prüfungen & UI-Hilfen
│   │   ├── tracking_service.py  # Einfaches Tracking (CSV)
│   │   └── org_structure_service.py # Organisationsstruktur (für DS)
│   ├── controllers/             # UI-Orchestrierung (ruft Services)
│   └── views/                   # UI-Darstellung
├── templates/                   # Word-Vorlagen
│   ├── MD_Ausblick.docx
│   ├── MD_Feedback.docx
│   ├── MD_Rückblick.docx
│   └── MD_Rückblick_Probezeit.docx
├── sap_stammdaten/              # SAP-Export für Stammdaten
│   └── EXPORT.xlsx
├── sap_massenupload/            # SAP-Export für Massenupload
├── tracking/                    # Tracking-Daten
│   └── versand/                 # Erstellte & versende MD-Dokumente
│   ├── ds_export/               # Data Science Export
│   └── logging/                 # Logging Dateien
├── ruecklauf/                   # Eingehende Dokumente
│   ├── unverarbeitet/           # Zielorder Mailanhänge 
|   |   └── manuell/             # Zielordner bei fehlerhafen Verarbeitung
│   └── verarbeitet/             # Zielordner bei erfolgreicher Verarbeitung
└── requirements.txt             # Python-Abhängigkeiten
```

## 🚀 Installation

1. **Python installieren**:
   Im Service Portal vom AFI Python 3.x bestellen (kostenlos)

2. **Pakete installieren** (einmalig):
   Doppelklick auf **Install.bat**
   - Installiert alle benötigten Python-Pakete
   - Dauert nur beim ersten Mal

3. **Anwendung starten**:
   Doppelklick auf **MD-App.bat**
   - Startet die Anwendung sofort
   - Keine Installation nötig

## 📖 Verwendung

### 1. SAP Stammdaten prüfen
- Lädt und validiert die EXPORT.xlsx
- Prüft Pflichtspalten und Datenqualität
- Zeigt auffällige Einträge (Duplikate, fehlende VG-PN, etc.)

### 2. MD-Versand
- **Massenversand**: Jahreslauf für alle Vorgesetzten
- **Einzelversand**: Unterjährige Einzelversendung
- **VG-MA-Verhältnis**: Neue Beziehungen anlegen

### 3. Maileingang verwalten
- Scannt Outlook-Postfach "VD-GS HR"
- Erkennt MD-Anhänge automatisch
- Speichert Anhänge nach `<root>/ruecklauf/unverarbeitet`


### 4. MD-Dokumente verarbeiten
- Verarbeitet DOCX/PDF aus `<root>/ruecklauf/unverarbeitet` (nur Top-Level)
- DOCX OK → `<root>/ruecklauf/verarbeitet`
- PDF Rückblick/Ausblick OK → RPA-Ziel (config `paths.output_dir` oder UI-Feld)
- PDF Feedback OK → `<root>/ruecklauf/feedbacks`
- Fehler/Prüfung nötig → `<root>/ruecklauf/unverarbeitet/manuell`
- Exporte: SAP-Massenupload, DS-Export

### 5. MD-Dashboard
- Status-Übersicht aller MD-Dokumente
- Filterung nach Manager, Status, Jahr
- Manuelle Status-Anpassungen
- CSV-Export

## ⚙️ Konfiguration

### config.yaml (Auszug)
```yaml
paths:
  base: ".."
  sap_stammdaten: "../sap_stammdaten/EXPORT.xlsx"
  templates:
    ausblick: "../templates/MD_Ausblick.docx"
    rueckblick: "../templates/MD_Rückblick.docx"
    rueckblick_probezeit: "../templates/MD_Rückblick_Probezeit.docx"
    feedback: "../templates/MD_Feedback.docx"
  ruecklauf:
    root: "../ruecklauf"
    unverarbeitet: "../ruecklauf/unverarbeitet"
    verarbeitet: "../ruecklauf/verarbeitet"
    manuell: "../ruecklauf/unverarbeitet/manuell"
    feedbacks: "../ruecklauf/feedbacks"
    logs_dir: "../ruecklauf/logs"
  tracking_dir: "../tracking"
  sap_massenupload: "../app/sap_massenupload/massenupload.xlsx"
  ds_export: "../app/tracking/ds_export/docx_extract.csv"
  rpa_input_dir: "K:/VD-GS-PUO-Personal/100 Roboter/Input"
  output_dir: "K:/VD-GS-PUO-Personal/100 Roboter/Input"

mail:
  send_mode: "display"  # "display" für Test, "send" für Produktiv
  from_address: ""
  bcc: ""
  # ...
```

## 🔧 Geschäftsregeln

### Dokumenttypen pro Mitarbeiter:
1. **Austritt Okt-Jan**: Nur Rückblick
2. **Probezeit Ende Okt-Jan**: Rückblick_Probezeit + Ausblick
3. **Probezeit Ende Jun-Sep**: Nur Ausblick
4. **Standard**: Rückblick + Ausblick

### E-Mail-Versand:
- **Jahreslauf**: Mit Feedback-Dokumenten
- **Unterjährig**: Ohne Feedback-Dokumente

## 📊 Tracking-System

Das System trackt den Status aller MD-Dokumente:
- **ausstehend**: Dokument erwartet
- **erhalten**: Dokument eingegangen
- **prüfung_nötig**: Manuelle Prüfung erforderlich
- **erübrigt**: Dokument nicht mehr benötigt

## 🛠️ Entwicklung

### Code-Struktur (vereinfacht):
- **core**: `data_loader.py`, `logging_config.py`, `constants.py`, `exceptions.py`, `utils.py`
- **adapters**: Outlook/Word/Docx Anbindung
- **services**: Geschäftslogik (keine UI)
- **controllers**: UI-Orchestrierung
- **views**: UI-Darstellung

## 🐛 Fehlerbehebung

### Häufige Probleme:
1. **Excel-Datei nicht gefunden**: Pfad in config.yaml prüfen
2. **Outlook-Verbindung**: Outlook muss installiert und konfiguriert sein
3. **Word-Vorlagen**: Alle Vorlagen müssen in templates/ vorhanden sein
4. **Berechtigungen**: Schreibzugriff auf alle Verzeichnisse erforderlich

