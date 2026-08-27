@echo off
setlocal
cd /d "C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS"
echo [%date% %time%] watchdog start>> output\_pdf_supervisor.log
:loop
python -u scripts\run_pdf_supervisor_loop.py
echo [%date% %time%] supervisor exited %ERRORLEVEL% - restart in 8s>> output\_pdf_supervisor.log
timeout /t 8 /nobreak >nul
goto loop
