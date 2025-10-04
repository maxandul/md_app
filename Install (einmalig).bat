@echo off
setlocal enabledelayedexpansion

echo ========================================
echo    MD-Prozess-Tool - Installation
echo ========================================
echo.
echo Diese Installation kann einige Minuten dauern.
echo Bitte warte, bis alle Schritte abgeschlossen sind.
echo.

REM Proxy-Einstellungen für Swisscom
set PROXY_URL=http://gateway.swisscom.zscloud.net:9400
set TRUSTED_HOSTS=--trusted-host pypi.org --trusted-host files.pythonhosted.org

echo Schritt 1: Python pruefen...
python --version >nul 2>&1
if errorlevel 1 (
    echo FEHLER: Python ist nicht installiert oder nicht im PATH verfügbar.
    echo Bitte bestelle Python im Serviceportal des AFI (kostenlose Standardapplikation). 
    echo.
    pause
    exit /b 1
)
python --version
echo Python gefunden! [OK]
echo.

echo Schritt 2: Pip aktualisieren (bitte warten)...
echo Versuche Pip-Update mit Proxy...
python -m pip install --upgrade pip %TRUSTED_HOSTS% --proxy %PROXY_URL% >nul 2>&1
if errorlevel 1 (
    echo WARNUNG: Pip-Update mit Proxy fehlgeschlagen. Versuche ohne Proxy...
    python -m pip install --upgrade pip %TRUSTED_HOSTS% >nul 2>&1
    if errorlevel 1 (
        echo WARNUNG: Pip konnte nicht aktualisiert werden. Versuche mit bestehender Version...
    ) else (
        echo Pip aktualisiert (ohne Proxy)! [OK]
    )
) else (
    echo Pip aktualisiert! [OK]
)
echo.

echo Schritt 3: Benötigte Pakete installieren (bitte warten)...
echo Dies kann einige Minuten dauern...
echo.

echo Versuche Installation mit Proxy...
pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org --proxy %PROXY_URL% -r requirements.txt
if errorlevel 1 (
    echo.
    echo WARNUNG: Installation mit Proxy fehlgeschlagen. Versuche ohne Proxy...
    pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt
    if errorlevel 1 (
        echo.
        echo FEHLER: Installation der Pakete fehlgeschlagen!
        echo Mögliche Ursachen:
        echo - Internet-Verbindung unterbrochen
        echo - Firewall blockiert die Verbindung
        echo - Proxy-Einstellungen fehlerhaft
        echo.
        echo Bitte kontaktiere den IT-Support.
        echo.
        pause
        exit /b 1
    ) else (
        echo Pakete erfolgreich installiert (ohne Proxy)! [OK]
    )
) else (
    echo Pakete erfolgreich installiert! [OK]
)

echo.
echo Schritt 4: Installation erfolgreich abgeschlossen! [OK]
echo.
echo ========================================
echo    Installation erfolgreich!
echo ========================================
echo.
echo Alle benötigten Pakete wurden installiert.
echo Du kannst jetzt die MD-App.bat starten.
pause
