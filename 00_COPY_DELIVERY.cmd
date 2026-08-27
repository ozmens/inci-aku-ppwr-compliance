@echo off
cd /d "%~dp0"
set "SRC=C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output"
set "DST=%~dp0delivery"
if not exist "%SRC%" (
  echo Kaynak yok: %SRC%
  exit /b 1
)
mkdir "%DST%" 2>nul
echo Tam kopya (~16 GB) — Render diski veya İnci sunucusu icin.
for %%D in (
  01_STARTER_INDIVIDUAL_DELIVERY_REV00
  02_INDUSTRIAL_DELIVERY_REV00
  03_CONTAINER_DELIVERY_REV00
  04_COMPONENT_SPARE_DELIVERY_REV00
) do (
  echo === %%D ===
  robocopy "%SRC%\%%D" "%DST%\%%D" /E /NFL /NDL /NJH /NJS /nc /ns /np
)
copy /Y "%SRC%\INCI_AKU_PPWR_STARTER_MASTER_Rev00.xlsx" "%DST%\" >nul
copy /Y "%SRC%\INCI_AKU_PPWR_INDUSTRIAL_MASTER_FROM_EXCEL_Rev00.xlsx" "%DST%\" >nul
echo Bitti.
