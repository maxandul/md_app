@echo off
setlocal
chcp 65001 >nul

echo =======================================================
echo   INSTALLIERE MD ENVIRONMENT (Benutzerkontext)
echo =======================================================
echo.

REM --- In das Verzeichnis der Batch wechseln ---
cd /d "%~dp0"

REM --- Proxy-Einstellungen (falls noetig) ---
set HTTP_PROXY=http://gateway.swisscom.zscloud.net:9400
set HTTPS_PROXY=http://gateway.swisscom.zscloud.net:9400

REM --- Prüfen, ob Python verfügbar ist ---
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ❌ Python wurde nicht gefunden.
    echo Bitte bestelle zuerst Python im Service Portal des AFI - kostenlose Standardapplikation
    pause
    exit /b 1
)

REM --- Python-Version anzeigen ---
for /f "delims=" %%v in ('python --version 2^>^&1') do set PY_VER=%%v
echo ✅ Gefundene Python-Version: %PY_VER%
echo.

REM --- Sicherstellen, dass pip verfuegbar ist ---
python -m ensurepip --default-pip >nul 2>nul
if %errorlevel% neq 0 (
    echo ❌ Konnte pip nicht initialisieren.
    echo Bitte fuehre Python einmal mit Adminrechten aus oder installiere pip manuell.
    pause
    exit /b 1
)

REM --- pip aktualisieren (nur Benutzerkontext, mit Proxy) ---
echo 📦 Aktualisiere pip...
python -m pip install --user --upgrade pip ^
    --proxy http://gateway.swisscom.zscloud.net:9400 ^
    --trusted-host pypi.org ^
    --trusted-host files.pythonhosted.org
if %errorlevel% neq 0 (
    echo ⚠️  Konnte pip nicht aktualisieren. Fahre trotzdem fort...
)

REM --- Requirements installieren ---
if not exist requirements.txt (
    echo ❌ Datei requirements.txt wurde nicht gefunden.
    echo Bitte lege sie in dasselbe Verzeichnis wie dieses Script.
    pause
    exit /b 1
)

echo 📥 Installiere Python-Dependencies aus requirements.txt ...
python -m pip install --user -r requirements.txt ^
    --proxy http://gateway.swisscom.zscloud.net:9400 ^
    --trusted-host pypi.org ^
    --trusted-host files.pythonhosted.org
if %errorlevel% neq 0 (
    echo ❌ Fehler beim Installieren der Dependencies.
    echo Bitte pruefe deine Internetverbindung oder Proxy-Einstellungen.
    pause
    exit /b 1
)

echo.
echo -------------------------------------------------------
echo ✅ Installation abgeschlossen! MD-App ist bereit.
echo -------------------------------------------------------
pause
endlocal
exit /b 0
