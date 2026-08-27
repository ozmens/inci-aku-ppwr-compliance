@echo off
REM Free stale PPWR ports (8790 zombie / old Vite)
cd /d "%~dp0"
echo Killing listeners on 8790 8791 5173...
for %%P in (8790 8791 5173) do (
  for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":%%P .*LISTENING"') do (
    echo   port %%P pid %%a
    taskkill /F /PID %%a /T >nul 2>&1
  )
)
echo Done. Simdi 00_START_PPWR_YAZILIMI.cmd calistirin.
pause
