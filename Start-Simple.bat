@echo off
echo ========================================
echo    MD-Prozess-Tool - Einfacher Start
echo ========================================
echo.

rem Arbeitsverzeichnis setzen
cd /d "%~dp0"
echo Arbeitsverzeichnis: %CD%

rem Python prüfen
python --version >nul 2>&1
if errorlevel 1 (
    echo FEHLER: Python nicht gefunden!
    pause
    exit /b 1
)

rem App starten
echo Starte MD-Prozess-Tool...
echo.
cd app
echo Neues Arbeitsverzeichnis: %CD%
echo.
python main.py

if errorlevel 1 (
    echo.
    echo FEHLER: App konnte nicht gestartet werden.
    echo.
    pause
) else (
    echo.
    echo App erfolgreich beendet.
)

pause
