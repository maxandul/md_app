# MD-Prozess-Tool - App-Struktur

## Architektur-Übersicht

Das MD-Prozess-Tool folgt einer sauberen MVC-Architektur mit Service-orientiertem Design:

```
app/
├── __init__.py              # App-Module-Export
├── main.py                  # Hauptanwendung (GUI)
├── constants.py             # Zentrale Konstanten
├── exceptions.py            # Exception-Klassen
├── logging_config.py        # Logging-Konfiguration
├── config.yaml             # Anwendungskonfiguration
├── data_loader.py          # Datenlade-Funktionen
├── utils.py                # Hilfsfunktionen
├── simple_tracking.py      # Tracking-System
├── controllers/            # Controller-Layer
│   ├── __init__.py
│   ├── dashboard_controller.py
│   ├── ruecklauf_controller.py
│   ├── stammdaten_controller.py
│   ├── verarbeitung_controller.py
│   └── versand_controller.py
├── services/               # Service-Layer (Geschäftslogik)
│   ├── __init__.py
│   ├── dashboard_service.py
│   ├── document_service.py
│   ├── email_service.py
│   ├── export_service.py
│   ├── file_service.py
│   ├── outlook_service.py
│   └── sap_data_service.py
└── views/                  # View-Layer (UI)
    ├── __init__.py
    ├── dashboard_view.py
    ├── ruecklauf_view.py
    ├── stammdaten_view.py
    ├── ui_utils.py
    ├── verarbeitung_view.py
    └── versand_view.py
```

## Design-Prinzipien

### 1. **Separation of Concerns**
- **Controllers**: Delegieren an Services, keine Geschäftslogik
- **Services**: Enthalten die gesamte Geschäftslogik
- **Views**: Nur UI-Aufbau, keine Logik

### 2. **Dependency Injection**
- Services werden über Controller aufgerufen
- Keine direkten Abhängigkeiten zwischen Views und Services

### 3. **Error Handling**
- Zentrale Exception-Klassen in `exceptions.py`
- Einheitliches Logging über `logging_config.py`
- Graceful Degradation bei Fehlern

### 4. **Konfiguration**
- Externe Konfiguration in `config.yaml`
- Zentrale Konstanten in `constants.py`
- Umgebungsabhängige Einstellungen

## Verwendung

```python
from app import App, MDConstants

# Anwendung starten
app = App()
app.mainloop()
```

## Entwicklung

### Neue Features hinzufügen:
1. **Service**: Geschäftslogik in `services/`
2. **Controller**: Delegation in `controllers/`
3. **View**: UI-Aufbau in `views/`

### Best Practices:
- Type Hints verwenden
- Docstrings für alle Funktionen
- Exception-Handling implementieren
- Logging für wichtige Operationen
