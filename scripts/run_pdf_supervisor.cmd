@echo off
setlocal
cd /d "%~dp0.."
echo [%DATE% %TIME%] PDF supervisor start> output\_pdf_supervisor.log
:loop
python -u scripts\pdf_one_batch.py --limit 24 >> output\_pdf_supervisor.log 2>&1
if errorlevel 1 (
  echo [%DATE% %TIME%] batch error — retry after wait>> output\_pdf_supervisor.log
  timeout /t 5 /nobreak >nul
)
findstr /C:"DONE" output\_pdf_supervisor.log >nul
if %ERRORLEVEL%==0 (
  echo [%DATE% %TIME%] all PDFs done>> output\_pdf_supervisor.log
  python -u scripts\finalize_product_level_engine.py >> output\_pdf_supervisor.log 2>&1
  goto end
)
REM stop if last 3 lines all say batch_ok=0 and no progress
timeout /t 2 /nobreak >nul
goto loop
:end
echo DONE
endlocal
