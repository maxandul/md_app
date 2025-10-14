# MD-Prozess-Tool

Ein Python-basiertes Tool zur Verwaltung des Mitarbeitenden-Dialog (MD) Prozesses.

## 📋 Übersicht

Das MD-Prozess-Tool automatisiert die Verwaltung des jährlichen Mitarbeitenden-Dialogs:

- **SAP Stammdaten-Validierung**: Prüfung und Validierung der Datei EXPORT.xlsx (SAP ad-hoc Query "VD_MD")
- **Dokumenten-Generierung**: Automatische Erstellung von MD-Dokumenten (Rückblick/Ausblick/Probezeit/Feedback Vorlage)
- **E-Mail-Versand**: Versendung der generierten Dokumente an Vorgesetzte (Massenversand & Einzelversand)
- **Rücklauf-Verarbeitung**: Automatische Verarbeitung eingehender MD-Dokumente aus Outlook
- **Tracking-System**: Vollständiges Tracking aller versendeten und eingegangenen Dokumente
- **Erinnerungsfunktion**: Automatische Erinnerungsmails an Vorgesetzte bei ausstehenden Dokumenten
- **Dashboard**: Übersicht über den kompletten MD-Prozess mit Filterfunktionen

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
│   │   ├── erinnerung_service.py # Erinnerungs-Funktionalität
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
│   ├── massenupload.xlsx        # Automatisch generierter SAP-Upload
│   └── archiv/                  # Archivierte Uploads
├── tracking/                    # Tracking-Daten & Logs
│   ├── md_logging_2024.csv      # Tracking MD-Durchlauf 2024
│   ├── md_logging_2025.csv      # Tracking MD-Durchlauf 2025
│   ├── app.log                  # Anwendungs-Log
│   ├── versand_2024/            # Versendete Dokumente 2024
│   │   └── VG_<PN>/             # Pro Vorgesetzten ein Ordner
│   ├── versand_2025/            # Versendete Dokumente 2025
│   │   └── VG_<PN>/             # Pro Vorgesetzten ein Ordner
│   ├── ds_export/               # Data Science Exports
│   │   ├── docx_extract_2024.csv # DOCX-Daten 2024
│   │   └── docx_extract_2025.csv # DOCX-Daten 2025
│   └── org_structure/           # Organisationsstruktur-Daten
├── ruecklauf/                   # Eingehende Dokumente
│   ├── unverarbeitet/           # Zielordner Mailanhänge 
│   │   └── manuell/             # Dokumente mit Fehlern/Prüfbedarf
│   ├── verarbeitet/             # Erfolgreich verarbeitete DOCX
│   ├── feedbacks/               # Feedback-PDFs (separate Ablage)
│   └── logs/                    # Verarbeitungs-Logs
│       └── processing_log.csv   # Detailliertes Verarbeitungsprotokoll
├── export/                      # CSV-Exporte aus Dashboard
├── requirements.txt             # Python-Abhängigkeiten
└── Start-Simple.bat             # Einfacher Start (ohne Optionen)
```

## 🚀 Installation & Start

1. **Python installieren**:
   Im Service Portal vom AFI Python 3.x bestellen (kostenlos)

2. **Pakete installieren** (einmalig):
   Doppelklick auf **Install (einmalig).bat**
   - Installiert alle benötigten Python-Pakete
   - Dauert nur beim ersten Mal (~2-3 Minuten)

3. **Anwendung starten**:
   Doppelklick auf **Start-Simple.bat**
   - Startet die Anwendung sofort
   - Keine weiteren Schritte nötig
   - Prüft Python-Installation automatisch

### Voraussetzungen
- Python 3.8 oder höher
- Microsoft Outlook (installiert und konfiguriert)
- Microsoft Word (für Vorlagen-Befüllung)
- Zugriff auf das Gruppenpostfach "VD-GS HR"

## 📖 Verwendung

### 🗓️ MD-Durchlaufjahr auswählen
**Oberhalb aller Tabs** findest du die Jahr-Auswahl für den aktiven MD-Durchlauf:

- **📅 MD-Durchlauf**: Wähle das Jahr des Rückblicks (z.B. 2025)
- Automatisch wird angezeigt: "Rückblick auf 2025 • Ausblick auf 2026"
- Alle Tabs verwenden automatisch dieses Jahr
- Tracking, Export und Versand werden jahr-spezifisch gespeichert

**Wann welches Jahr wählen?**
- Oktober-Dezember 2025 → Jahr 2025 (MD-Start)
- Januar-April 2026 → Jahr 2025 (Nachläufer-Phase)
- Mai-September 2026 → Jahr 2026 (Vorbereitung neuer Durchlauf)

### 1️⃣ SAP Stammdaten prüfen
**Zweck**: Validierung der SAP-Stammdaten vor dem MD-Versand

- Lädt und validiert die Datei `sap_stammdaten/EXPORT.xlsx`
- Prüft Pflichtspalten und Datenqualität
- Zeigt auffällige Einträge:
  - Duplikate (mehrfach vorkommende Personalnummern)
  - Fehlende Vorgesetzten-PN
  - BsGrd = 0 (noch nicht zugeordnet)
- **Info-Button** ⓘ zeigt detaillierte Anleitung zur EXPORT.xlsx-Vorbereitung

### 2️⃣ MD-Versand

#### Massenversand (Jährlicher Durchlauf)
- Vorgesetzte auswählen (Mehrfachauswahl möglich)
- **Vorschau**: Zeigt alle zu versendenden E-Mails
- **Versand-Modi**:
  - "Generieren & Versenden": Sofortiger Versand
  - "Als Entwurf speichern": Outlook-Entwürfe zur manuellen Prüfung
- **Automatische Dokumenttyp-Auswahl** pro Mitarbeiter:
  - Standard: Rückblick + Ausblick
  - Austritt Okt-Jan: Nur Rückblick
  - Probezeit-Ende Okt-Jan: Rückblick Probezeit + Ausblick
  - Probezeit-Ende Jun-Sep: Nur Ausblick
- Pro Vorgesetzten wird eine **Feedback-Vorlage** erstellt
- Alle Dokumente werden im **Tracking-System erfasst** (außer Probezeit-Rückblick)
- **Sicherheitsabfrage**: Vor dem Versand erscheint ein Dialog mit dem aktiven MD-Durchlaufjahr zur Bestätigung

#### Einzelversand (Unterjährig)
- **Einen** Vorgesetzten auswählen
- Gewünschte Mitarbeitende auswählen
- **Manuelle Dokumenttyp-Auswahl**:
  - ☑ Rückblick
  - ☑ Ausblick
  - ☑ Rückblick Probezeit
- **Kein** Feedback-Dokument im Einzelversand
- **Vorschau** und Versand-Modi wie beim Massenversand
- **Sicherheitsabfrage**: Vor dem Versand erscheint ein Dialog zur Bestätigung

#### VG-MA-Verhältnis anlegen
- Neue Vorgesetzten-Mitarbeiter-Beziehung erstellen
- Nützlich bei:
  - Neueinstellungen
  - Abteilungswechsel
  - Vertretungsregelungen
- Erstellt neuen Datensatz in EXPORT.xlsx
- **Warnung**: Änderungen sofort gespeichert, App-Neustart erforderlich!

### 3️⃣ Maileingang verwalten (Rücklauf)
**Zweck**: Eingehende MD-Dokumente aus Outlook automatisch abholen

- Scannt Outlook-Gruppenpostfach "VD-GS HR"
- **Dokumenttyp-Erkennung**:
  - Rückblick (Word/PDF)
  - Ausblick (Word/PDF)
  - Feedback (PDF)
  - Probezeit-Rückblick
  - Sonstige Dateien
- **Drei Kategorien** (alle auf einen Blick):
  - ✓ **Kopiert & verschoben**: Nur MD-Anhänge → Mail verschoben nach "12 Mitarbeitenden-Dialog"
  - ⚠ **Prüfen erforderlich**: Fremde/Probezeit-Anhänge → Mail bleibt im Posteingang
  - ○ **Übersprungen**: Keine MD-Anhänge → Mail unverändert
- Anhänge werden nach `ruecklauf/unverarbeitet` kopiert

### 4️⃣ MD-Dokumente verarbeiten
**Zweck**: Eingegangene Dokumente validieren und für SAP vorbereiten

**Drei-Schritte-Prozess**:

1. **DOCX prüfen**:
   - Liest Word-Dokumente (Rückblick/Ausblick)
   - Validiert Pflichtfelder (Name, PN, Gesamteindruck)
   - Prüft gegen SAP-Stammdaten
   - **Aktualisiert Tracking-System**
   
2. **Export & Verschieben**:
   - Erstellt **SAP-Massenupload** (`sap_massenupload/massenupload.xlsx`)
   - Erstellt **DataScience-Export** (`tracking/ds_export/docx_extract_{jahr}.csv`)
   - DOCX "ok" → `ruecklauf/verarbeitet`
   - DOCX "manuell" → `ruecklauf/unverarbeitet/manuell`
   
3. **PDFs verarbeiten**:
   - Erkennt Dokumenttyp aus Dateinamen (Deutsche Umlaute werden korrekt erkannt!)
   - Extrahiert Personalnummer (6-stellig)
   - **Aktualisiert Tracking-System**
   - Verteilung:
     - Feedback → `ruecklauf/feedbacks/`
     - Rückblick/Ausblick → RPA-Zielordner (für Roboter-Upload nach SAP)
     - Fehler → `ruecklauf/unverarbeitet/manuell`

**Einstellungen**:
- Batchgröße (begrenzt Anzahl Dateien pro Lauf)
- RPA-Zielordner (wohin PDFs für SAP-Upload verschoben werden)

**Hinweis**: Das Durchlauf-Jahr wird oben im globalen MD-Durchlauf-Feld festgelegt.

**Sicherheitsabfrage**: Vor dem Start der Verarbeitung erscheint ein Dialog mit dem aktiven MD-Durchlaufjahr zur Bestätigung.

### 5️⃣ MD-Dashboard
**Zweck**: Status-Übersicht & Erinnerungen

**Funktionen**:
- **Aktualisieren**: Lädt Tracking-Daten aus `tracking/md_logging_{jahr}.csv`
  (Jahr entspricht dem oben ausgewählten MD-Durchlauf)
- **Filter**:
  - Nach Namen (Vorgesetzten/Mitarbeiter)
  - Nach Status (ausstehend/erhalten/prüfung_nötig/erübrigt)
- **Manuelle Anpassung**: Status/Grund einzelner Einträge korrigieren
- **Export CSV**: Gefilterte Ansicht exportieren
- **Erinnerungen senden**:
  1. Zeilen auswählen (Mehrfachauswahl mit Strg/Shift)
  2. "Vorschau generieren" → Zeigt alle geplanten Erinnerungsmails
  3. "Erinnerung versenden" oder "Als Entwurf speichern"
  - Mails werden **pro Vorgesetzten gruppiert**
  - Enthält nur **ausstehende** Dokumente der Auswahl

**Status-Bedeutung**:
- 🔵 **ausstehend**: Dokument noch nicht eingegangen
- ✅ **erhalten**: Vollständig eingegangen (Word + PDF)
- ⚠️ **prüfung_nötig**: Fehler bei Verarbeitung (siehe Grund-Spalte)
- ⭕ **erübrigt**: Manuell als nicht mehr relevant markiert

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
  # DS-Export wird automatisch jahr-spezifisch erstellt (docx_extract_{jahr}.csv)
  rpa_input_dir: "K:/VD-GS-PUO-Personal/100 Roboter/Input"
  output_dir: "K:/VD-GS-PUO-Personal/100 Roboter/Input"

mail:
  send_mode: "display"  # "display" für Test, "send" für Produktiv
  from_address: ""
  bcc: ""
  # ...
```

## 🔧 Geschäftsregeln

### Automatische Dokumenttyp-Auswahl (Massenversand):
Die Anwendung wählt automatisch die richtigen Dokumenttypen basierend auf Mitarbeiter-Status:

1. **Austritt zwischen Oktober und Januar**: Nur Rückblick
   - Grund: Mitarbeiter verlässt Organisation, kein Ausblick nötig

2. **Probezeit-Ende zwischen Oktober und Januar**: Rückblick Probezeit + Ausblick
   - Grund: Spezielle Probezeit-Reflexion + Planung für Folgejahr

3. **Probezeit-Ende zwischen Juni und September**: Nur Ausblick
   - Grund: Probezeit bereits abgeschlossen, Fokus auf Zukunft

4. **Standard (alle anderen)**: Rückblick + Ausblick
   - Grund: Regulärer MD-Zyklus

### Tracking-Besonderheiten:
- ✅ **Rückblick**: Word + PDF werden getrackt (erwartet: 2 Dokumente)
- ✅ **Ausblick**: Word + PDF werden getrackt (erwartet: 2 Dokumente)
- ⚠️ **Rückblick Probezeit**: Wird NICHT getrackt (separater Prozess)
- ✅ **Feedback**: Nur PDF wird getrackt (pro Vorgesetzten)

### E-Mail-Versand:
- **Massenversand (Jahreslauf)**: Mit Feedback-Vorlage für Vorgesetzten
- **Einzelversand (Unterjährig)**: Ohne Feedback-Vorlage

### Duplikat-Erkennung:
- **Word-Dokumente**: Eindeutige Zuordnung über VG-PN + MA-PN + Dokumenttyp
- **PDF-Dokumente**: Nur MA-PN bekannt → Bei Mehrfachanstellung manuelle Prüfung nötig

## 📊 Tracking-System

Das System trackt vollständig alle MD-Dokumente in jahr-spezifischen Dateien (`tracking/md_logging_{jahr}.csv`).

**Jahr-basierte Dateistruktur**:
- Jeder MD-Durchlauf wird separat gespeichert
- Beispiel: `md_logging_2025.csv` enthält alle Daten für den MD-Durchlauf 2025
- Ermöglicht saubere Trennung zwischen Jahren
- Alte Durchläufe können archiviert werden

### Erfasste Daten pro Eintrag:
- `log_id`: Eindeutige ID
- `vg_pn`, `vg_name`: Vorgesetzten-Info
- `ma_pn`, `ma_name`: Mitarbeiter-Info
- `doc_type`: Dokumenttyp (z.B. "Rückblick Word", "Ausblick PDF")
- `erwartet`: Anzahl erwarteter Dokumente (meist 1)
- `erhalten`: Anzahl eingegangener Dokumente
- `status`: Aktueller Status
- `status_grund`: Fehlermeldung bei Problemen
- `versendet_am`: Zeitpunkt des Versands
- `zuletzt_erinnert_am`: Zeitpunkt der letzten Erinnerung

### Status-Übergänge:
1. **Versand** → Status: `ausstehend` (erwartet=1, erhalten=0)
2. **Dokument eingeht** → Status: `erhalten` (erhalten=1)
3. **Fehler bei Verarbeitung** → Status: `prüfung_nötig` + Fehlergrund
4. **Manuell korrigiert** → Status: `erübrigt` oder zurück zu `ausstehend`

### Automatisches Update:
- **Bei Versand**: Neue Einträge werden erstellt
- **Bei DOCX-Verarbeitung**: Word-Dokumente als erhalten markiert
- **Bei PDF-Verarbeitung**: PDF-Dokumente als erhalten markiert
- **Bei Fehlern**: Status und Fehlergrund automatisch gesetzt

## 🛠️ Entwicklung

### Code-Struktur (MVC-Pattern):
- **core**: `data_loader.py`, `logging_config.py`, `constants.py`, `exceptions.py`, `utils.py`
- **adapters**: Technische Anbindungen (Outlook, Word, DOCX)
  - `mail_outlook.py`: E-Mail-Versand via Outlook COM
  - `word_outlook.py`: Word-Template-Befüllung
  - `docx_reader.py`: Lesen von Content-Controls aus DOCX
- **services**: Geschäftslogik (keine UI-Abhängigkeiten)
  - `document_service.py`: DOCX/PDF-Verarbeitung, Validierung
  - `email_service.py`: Versand-Flows (Massen-/Einzelversand)
  - `dispatch_service.py`: Dokument-Generierung
  - `tracking_service.py`: Tracking-System (CSV-basiert)
  - `export_service.py`: SAP-/DataScience-Exporte
  - `erinnerung_service.py`: Erinnerungs-Funktionalität
  - `outlook_service.py`: Postfach-Scanning
- **controllers**: UI-Orchestrierung (verbindet Views & Services)
  - Jeder Tab hat einen eigenen Controller
  - Behandelt Button-Klicks, validiert Eingaben
  - Ruft Services auf und aktualisiert UI
- **views**: UI-Darstellung (Tkinter)
  - Erzeugt nur die UI-Elemente
  - Keine Geschäftslogik

### Wichtige Design-Entscheidungen:
- **Keine direkte UI-Service-Kopplung**: Services kennen keine UI
- **Tracking per CSV**: Einfach, lesbar, Excel-kompatibel
- **Outlook COM-Interface**: Native Outlook-Integration (kein IMAP)
- **Word-Templates**: Content-Controls für strukturierte Datenbefüllung

## 🐛 Fehlerbehebung

### Häufige Probleme:

#### 1. "Python nicht gefunden"
- **Lösung**: Python 3.8+ im Service Portal bestellen
- **Prüfung**: `python --version` in CMD ausführen

#### 2. "Excel-Datei EXPORT.xlsx nicht gefunden"
- **Lösung**: Datei in `sap_stammdaten/EXPORT.xlsx` ablegen
- **Prüfung**: Pfad in `config.yaml` unter `paths.sap_stammdaten` prüfen

#### 3. "Outlook-Verbindung fehlgeschlagen"
- **Lösung**: Outlook muss installiert, konfiguriert und geöffnet sein
- **Prüfung**: Outlook manuell starten und Postfach "VD-GS HR" öffnen
- **Hinweis**: Windows-Sicherheitswarnung "Programm greift auf Outlook zu" mit "Zulassen" bestätigen

#### 4. "Word-Vorlage nicht gefunden"
- **Lösung**: Alle Vorlagen in `templates/` ablegen:
  - `MD_Ausblick.docx`
  - `MD_Rückblick.docx`
  - `MD_Rückblick_Probezeit.docx`
  - `MD_Feedback.docx`
- **Wichtig**: Vorlagen müssen Content-Controls enthalten!

#### 5. "Zugriff verweigert" beim Speichern
- **Lösung**: Schreibzugriff auf alle Verzeichnisse prüfen
- **Typische Ordner**: `ruecklauf/`, `tracking/`, `sap_massenupload/`
- **Prüfung**: Als Administrator ausführen oder Berechtigungen anpassen

#### 6. "Tracking zeigt keine Daten"
- **Lösung**: `tracking/md_logging_{jahr}.csv` existiert erst nach erstem Versand für dieses Jahr
- **Prüfung**: Mindestens einmal Dokumente für das ausgewählte Jahr versenden
- **Hinweis**: Stelle sicher, dass das richtige Jahr oben ausgewählt ist

#### 7. "PDF wird nicht erkannt" / "Dokumenttyp: Unbekannt"
- **Ursache**: Dateiname enthält keine MD-Keywords
- **Lösung**: Datei umbenennen mit "Rückblick", "Ausblick" oder "Feedback" im Namen
- **Beispiel**: `Rückblick_2025_Max_Mustermann_123456.pdf`
- **Hinweis**: Deutsche Umlaute (ü, ä, ö) werden automatisch erkannt!

#### 8. "Mehrfachanstellung erkannt" bei PDF
- **Ursache**: Mitarbeiter hat mehrere Anstellungen bei verschiedenen VG
- **Lösung**: Manuelle Zuordnung erforderlich
- **Vorgehen**: 
  1. PDF in `ruecklauf/unverarbeitet/manuell` prüfen
  2. Im Dashboard den korrekten Eintrag finden
  3. Status manuell auf "erhalten" setzen

### Logging:
- **Hauptlog**: `tracking/app.log` (alle Anwendungs-Events)
- **Verarbeitungslog**: `ruecklauf/logs/processing_log.csv` (DOCX/PDF-Verarbeitung)
- Bei Problemen: Log-Dateien auf Fehlermeldungen prüfen

### Support:
- **Info-Buttons** ⓘ in jedem Tab erklären die Funktionen
- **Hover-Tooltips**: Maus über Buttons halten für zusätzliche Infos
- Bei technischen Problemen: Log-Dateien bereithalten
