@echo off
setlocal enabledelayedexpansion

echo ========================================
echo    MD-Prozess-Tool - DEBUG
echo ========================================
echo.
echo Diese Datei hilft bei der Fehlerdiagnose.
echo.

echo Schritt 1: System-Informationen
echo Arbeitsverzeichnis: %CD%
echo Batch-Datei-Pfad: %~dp0
echo.

echo Schritt 2: Python-Informationen
python --version
echo Python-Pfad:
where python
echo.

echo Schritt 3: Python-Module prüfen
echo Prüfe pandas...
python -c "import pandas; print('pandas OK')" 2>nul || echo pandas FEHLER
echo Prüfe tkinter...
python -c "import tkinter; print('tkinter OK')" 2>nul || echo tkinter FEHLER
echo Prüfe win32com...
python -c "import win32com.client; print('win32com OK')" 2>nul || echo win32com FEHLER
echo Prüfe yaml...
python -c "import yaml; print('yaml OK')" 2>nul || echo yaml FEHLER
echo.

echo Schritt 4: Projektstruktur prüfen
if exist "app" (
    echo app-Ordner gefunden: OK
    if exist "app\main.py" (
        echo app\main.py gefunden: OK
    ) else (
        echo app\main.py NICHT gefunden: FEHLER
    )
) else (
    echo app-Ordner NICHT gefunden: FEHLER
)
echo.

echo Schritt 5: Python-Modul-Import testen
echo Teste: python -m app.main
cd /d "%~dp0"
python -m app.main --help 2>nul || echo Modul-Import FEHLER
echo.

echo ========================================
echo    Debug-Informationen abgeschlossen
echo ========================================
echo.
echo Falls Fehler angezeigt wurden, kontaktiere den IT-Support.
echo.
pause
