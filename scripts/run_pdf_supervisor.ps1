$ErrorActionPreference = "Continue"
Set-Location "C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS"
$log = "output\_pdf_supervisor.log"
"[$([DateTime]::Now.ToString('s'))] supervisor start" | Out-File $log -Encoding utf8
$idle = 0
while ($true) {
  $out = & python -u scripts\pdf_one_batch.py --limit 24 2>&1
  $out | Tee-Object -FilePath $log -Append
  $text = ($out | Out-String)
  if ($text -match "DONE") {
    "[$([DateTime]::Now.ToString('s'))] all pdfs complete" | Tee-Object -FilePath $log -Append
    & python -u scripts\finalize_product_level_engine.py 2>&1 | Tee-Object -FilePath $log -Append
    break
  }
  if ($text -match "batch_ok=0") {
    $idle++
  } else {
    $idle = 0
  }
  if ($idle -ge 8) {
    "[$([DateTime]::Now.ToString('s'))] stopping after idle batches" | Tee-Object -FilePath $log -Append
    & python -u scripts\finalize_product_level_engine.py 2>&1 | Tee-Object -FilePath $log -Append
    break
  }
  Start-Sleep -Seconds 2
}
"[$([DateTime]::Now.ToString('s'))] supervisor end" | Tee-Object -FilePath $log -Append
