@echo off
echo ========================================
echo    MD-Prozess-Tool
echo ========================================
echo.

echo Schritt 1: Python pruefen (bitte warten)
python --version
echo.

echo Schritt 2: Abhaengigkeiten pruefen (bitte warten)
python -c "import pandas, tkinter, win32com.client, yaml, openpyxl"
echo.

echo Schritt 3: App starten (dieses Fenster nicht schliessen)
cd /d "%~dp0"
python -m app.main
echo.

echo Schritt 4: App beendet
echo Fenster kann geschlossen werden
pause
