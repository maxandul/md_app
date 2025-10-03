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
    echo Bitte führe zuerst 'Install (einmalig).bat' aus.
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
cd /d "%~dp0"
python -m app.main
if errorlevel 1 (
    echo.
    echo FEHLER: Die App ist mit einem Fehler beendet worden.
    echo Falls das Problem weiterhin besteht, kontaktiere den IT-Support.
    echo.
) else (
    echo.
    echo MD-Prozess-Tool wurde erfolgreich beendet. [OK]
    echo.
)

echo Schritt 4: Vorgang abgeschlossen
echo Du kannst dieses Fenster jetzt schließen.
pause
endlocal
