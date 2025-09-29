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
├── app/                          # Hauptanwendung
│   ├── main.py                   # GUI-Hauptanwendung (Tkinter)
│   ├── data_loader.py           # SAP-Datenlade und -verarbeitung
│   ├── dispatch.py              # Dokumentenversand und -generierung
│   ├── doc_processing.py        # Dokumentenverarbeitung (DOCX/PDF)
│   ├── docx_tools.py            # Word-Dokument-Tools
│   ├── word_tools.py            # Word-Template-Verarbeitung
│   ├── mail_send.py             # E-Mail-Versand (Outlook)
│   ├── simple_tracking.py       # Status-Tracking-System
│   ├── utils.py                 # Hilfsfunktionen
│   └── config.yaml              # Konfigurationsdatei
├── templates/                    # Word-Vorlagen
│   ├── MD_Ausblick.docx
│   ├── MD_Feedback.docx
│   ├── MD_Rückblick.docx
│   └── MD_Rückblick_Probezeit.docx
├── sap_stammdaten/              # SAP-Export-Dateien
│   └── EXPORT.xlsx
├── tracking/                    # Tracking-Daten
│   └── versand/
├── ruecklauf/                   # Eingehende Dokumente
│   ├── unverarbeitet/
│   ├── verarbeitet/
│   └── manuell/
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
- Verschiebt E-Mails nach Verarbeitung

### 4. MD-Dokumente verarbeiten
- Verarbeitet DOCX-Dokumente aus `ruecklauf/unverarbeitet/`
- Extrahiert Daten und validiert gegen SAP
- Exportiert für SAP-Massenupload und DS-System
- Verarbeitet PDF-Dokumente

### 5. MD-Dashboard
- Status-Übersicht aller MD-Dokumente
- Filterung nach Manager, Status, Jahr
- Manuelle Status-Anpassungen
- CSV-Export

## ⚙️ Konfiguration

### config.yaml
```yaml
paths:
  sap_stammdaten: "../sap_stammdaten/EXPORT.xlsx"
  templates:
    ausblick: "../templates/MD_Ausblick.docx"
    rueckblick: "../templates/MD_Rückblick.docx"
    # ...

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

### Code-Struktur:
- **main.py**: GUI-Hauptanwendung (Tkinter)
- **data_loader.py**: SAP-Datenverarbeitung
- **dispatch.py**: Dokumentenversand-Logik
- **doc_processing.py**: Dokumentenverarbeitung
- **simple_tracking.py**: Status-Tracking

## 🐛 Fehlerbehebung

### Häufige Probleme:
1. **Excel-Datei nicht gefunden**: Pfad in config.yaml prüfen
2. **Outlook-Verbindung**: Outlook muss installiert und konfiguriert sein
3. **Word-Vorlagen**: Alle Vorlagen müssen in templates/ vorhanden sein
4. **Berechtigungen**: Schreibzugriff auf alle Verzeichnisse erforderlich

