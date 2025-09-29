@echo off
echo ========================================
echo    MD-Prozess-Tool - Installation
echo ========================================
echo.

echo Schritt 1: Python pruefen
python --version
echo.

echo Schritt 2: Pip aktualisieren (bitte warten)
python -m pip install --upgrade pip
echo.

echo Schritt 3: Pakete installieren (bitte warten)
pip install -r requirements.txt
echo.

echo Schritt 4: Installation abgeschlossen
echo.
echo Du kannst jetzt MD-App-New.bat starten
echo.
pause
