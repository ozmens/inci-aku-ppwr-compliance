@echo off
cd /d "%~dp0"
echo [0/2] Eski portlari temizle (8791 / 5173)...
for %%P in (8791 5173) do (
  for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":%%P .*LISTENING"') do (
    taskkill /F /PID %%a /T >nul 2>&1
  )
)
timeout /t 1 /nobreak >nul

echo [1/2] Backend API (8791)...
start "INCI-PPWR-API" cmd /k "cd /d "%~dp0backend" && set INCI_PPWR_WEB=1&& set INCI_PPWR_VERSION=1.0.0&& set INCI_PPWR_ADMIN_USER=admin&& set INCI_PPWR_ADMIN_PASSWORD=160616&& python -m pip install -q -r requirements.txt && python -m uvicorn main:app --host 127.0.0.1 --port 8791 --reload"
timeout /t 2 /nobreak >nul
echo [2/2] Frontend (5173)...
cd /d "%~dp0app"
if not exist node_modules (
  call npm install
)
start "INCI-PPWR-UI" cmd /k "npm run dev"
echo.
echo Ac: http://localhost:5173
echo API: http://127.0.0.1:8791/docs
echo Login: admin / 160616
echo Stuck port: 00_KILL_STALE_PORTS.cmd
echo Delivery setleri SALT OKUNUR.
