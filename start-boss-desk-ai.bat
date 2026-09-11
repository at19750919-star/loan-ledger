@echo off
cd /d "%~dp0"
set PORT=3218
set OPEN_BROWSER=1
set "CHROME_PATH=C:\Program Files\Google\Chrome\Application\chrome.exe"
set "CHROME_PROFILE_DIRECTORY=Profile 1"

REM If the dashboard is already running, only open it instead of starting a second server.
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $r = Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:3218/' -TimeoutSec 2; if ($r.StatusCode -eq 200) { exit 0 } } catch {}; exit 1" >nul 2>&1
if not errorlevel 1 (
  start "" "%CHROME_PATH%" --profile-directory="%CHROME_PROFILE_DIRECTORY%" "http://127.0.0.1:3218/"
  exit /b 0
)

node server.js
if errorlevel 1 pause
