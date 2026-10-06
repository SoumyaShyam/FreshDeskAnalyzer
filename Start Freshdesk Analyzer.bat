@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is required. Install Node.js 18 or newer, then try again.
  pause
  exit /b 1
)

set "APP_URL=http://127.0.0.1:8001/"
set "HEALTH_URL=http://127.0.0.1:8001/api/default-folder"
powershell.exe -NoProfile -Command "try { Invoke-RestMethod -Uri '%HEALTH_URL%' -TimeoutSec 2 | Out-Null; exit 0 } catch { exit 1 }"
if not errorlevel 1 goto open_app

start "Freshdesk Snapshot Analyzer" /min cmd /c "node server.js"
timeout /t 2 /nobreak >nul
powershell.exe -NoProfile -Command "try { Invoke-RestMethod -Uri '%HEALTH_URL%' -TimeoutSec 2 | Out-Null; exit 0 } catch { exit 1 }"
if errorlevel 1 (
  echo The local server did not start. Check that port 8001 is available.
  pause
  exit /b 1
)

:open_app
start "" "%APP_URL%"
exit /b 0