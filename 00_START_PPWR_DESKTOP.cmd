@echo off
cd /d "%~dp0"
echo ============================================================
echo  Inci Aku PPWR Compliance Suite v1.0.0 — Desktop
echo  Frozen deliveries: READ-ONLY
echo ============================================================

echo [0/3] Port temizligi (8791 / 5173)...
for %%P in (8791 5173) do (
  for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":%%P .*LISTENING"') do (
    taskkill /F /PID %%a /T >nul 2>&1
  )
)
timeout /t 1 /nobreak >nul

echo [1/3] Backend API (8791)...
start "INCI-PPWR-API" cmd /k "cd /d "%~dp0backend" && set INCI_PPWR_WEB=0&& set INCI_PPWR_VERSION=1.0.0&& set INCI_PPWR_ADMIN_USER=admin&& set INCI_PPWR_ADMIN_PASSWORD=160616&& python -m pip install -q -r requirements.txt && python -m uvicorn main:app --host 127.0.0.1 --port 8791 --reload"
timeout /t 2 /nobreak >nul

echo [2/3] Frontend Vite (5173)...
cd /d "%~dp0app"
if not exist node_modules call npm install
start "INCI-PPWR-UI" cmd /k "npm run dev"
timeout /t 2 /nobreak >nul

echo [3/3] Electron kabugu...
cd /d "%~dp0desktop"
if not exist node_modules call npm install
echo Bekleniyor: API + UI hazir olunca pencere acilir...
echo Login: admin / 160616
call npm start
echo.
echo Desktop kapandi. API/UI pencereleri aciksa calismaya devam eder.
