@echo off
setlocal enabledelayedexpansion

echo ========================================
echo    MD-Prozess-Tool
echo ========================================
echo.
echo Starte MD-Prozess-Tool...
echo.

rem Schritt 1: Python pruefen
echo Schritt 1: Python pruefen...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo FEHLER: Python ist nicht installiert oder nicht im PATH verfügbar.
    echo Bitte bestelle Python im Serviceportal des AFI (kostenlose Standardapplikation).
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYV=%%v
echo Python gefunden: %PYV% [OK]
echo.

rem Schritt 2: Abhaengigkeiten pruefen
echo Schritt 2: Abhängigkeiten pruefen...
python -c "import pandas, tkinter, win32com.client, yaml, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo.
    echo FEHLER: Erforderliche Python-Pakete sind nicht installiert.
    echo.
    echo Lösung:
    echo 1. Führe 'Install (einmalig).bat' aus (falls noch nicht geschehen)
    echo 2. Oder installiere manuell: pip install -r requirements.txt
    echo.
    pause
    exit /b 2
)
echo Alle Abhängigkeiten gefunden! [OK]
echo.

rem Schritt 3: App starten
echo Schritt 3: MD-Prozess-Tool starten...
echo.
echo WICHTIG: Dieses Fenster nicht schließen, während die App läuft!
echo.

rem Arbeitsverzeichnis auf Batch-Datei-Pfad setzen
cd /d "%~dp0"
echo Arbeitsverzeichnis: %CD%

rem Prüfe ob app-Ordner existiert
if not exist "app" (
    echo FEHLER: app-Ordner nicht gefunden!
    echo Stelle sicher, dass die Batch-Datei im Projektverzeichnis liegt.
    echo.
    pause
    exit /b 3
)

rem App starten
echo Starte Python-Modul...
echo Versuche Modul-Start...
python -m app.main
if errorlevel 1 (
    echo.
    echo Modul-Start fehlgeschlagen, versuche direkten Start...
    cd app
    python main.py
    if errorlevel 1 (
        echo.
        echo FEHLER: Die App ist mit einem Fehler beendet worden.
        echo.
        echo Mögliche Lösungen:
        echo 1. Stelle sicher, dass alle Abhängigkeiten installiert sind
        echo 2. Führe 'Install (einmalig).bat' erneut aus
        echo 3. Kontaktiere den IT-Support bei anhaltenden Problemen
        echo.
        echo Drücke eine beliebige Taste zum Beenden...
        pause
        exit /b 3
    ) else (
        echo.
        echo MD-Prozess-Tool wurde erfolgreich beendet. [OK]
        echo.
    )
) else (
    echo.
    echo MD-Prozess-Tool wurde erfolgreich beendet. [OK]
    echo.
)

echo.
echo ========================================
echo Vorgang abgeschlossen
echo ========================================
echo.
echo Die MD-App wurde beendet.
echo Du kannst dieses Fenster jetzt schließen.
echo.
pause
