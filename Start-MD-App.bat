@echo off
setlocal enabledelayedexpansion

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
echo Setze Python-Pfad für direkte Imports...
set PYTHONPATH=%CD%\app;%CD%
echo PYTHONPATH: %PYTHONPATH%
echo.
cd app
echo Arbeitsverzeichnis: %CD%
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
