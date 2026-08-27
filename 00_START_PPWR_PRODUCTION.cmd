@echo off
cd /d "%~dp0"
echo === Inci Aku PPWR Compliance Suite v1.0.0 — PRODUCTION ===
echo.

set INCI_PPWR_WEB=1
set INCI_PPWR_VERSION=1.0.0
if not defined INCI_PPWR_ADMIN_USER set INCI_PPWR_ADMIN_USER=admin
if not defined INCI_PPWR_ADMIN_PASSWORD set INCI_PPWR_ADMIN_PASSWORD=160616

echo [0/3] Port 8791 temizle...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8791 .*LISTENING"') do (
  taskkill /F /PID %%a /T >nul 2>&1
)
timeout /t 1 /nobreak >nul

echo [1/3] UI build...
cd /d "%~dp0app"
if not exist node_modules call npm install
call npm run build
if errorlevel 1 (
  echo UI build FAILED
  exit /b 1
)

echo [2/3] Python deps...
cd /d "%~dp0backend"
python -m pip install -q -r requirements.txt

echo [3/3] Serving UI+API on http://127.0.0.1:8791
echo Login: admin / 160616
echo OPEN WORD / OPEN PDF = browser download  ·  ZIP = download
echo.
python -m uvicorn main:app --host 127.0.0.1 --port 8791
