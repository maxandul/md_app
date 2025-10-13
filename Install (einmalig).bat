@echo off
echo =======================================================
echo Installing MD Environment (user context, via Python)
echo =======================================================

REM --- Proxy settings (falls noetig anpassen oder auskommentieren) ---
set HTTP_PROXY=http://gateway.swisscom.zscloud.net:9400
set HTTPS_PROXY=http://gateway.swisscom.zscloud.net:9400

REM --- Install using the same interpreter as in IDE (user mode) ---
python -m pip install --user --upgrade pip
python -m pip install --user --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt

echo -------------------------------------------------------
echo Installation completed! Ready for MD-App.
echo -------------------------------------------------------
pause
