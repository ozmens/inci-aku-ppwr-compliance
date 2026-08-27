$ErrorActionPreference = "Continue"
$sets = "C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\INCI_AKU_PPWR_STARTER_PRODUCT_LEVEL_CUSTOMER_DELIVERY_REV00_CANDIDATE\01_PRODUCT_DOCUMENT_SETS"
$jobs = @(
  "1014048\04_Shipment_Statement.docx",
  "1014616\03_Label.docx",
  "1014631\02_EU_DoC.docx",
  "1014850\01_Technical_File.docx",
  "1015016\01_Technical_File.docx",
  "1015018\04_Shipment_Statement.docx",
  "1015132\02_EU_DoC.docx",
  "1015132\03_Label.docx",
  "1015334\03_Label.docx",
  "1015336\02_EU_DoC.docx"
)

function Kill-Word {
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
  Start-Sleep -Seconds 2
}

$ok = 0; $fail = 0
foreach ($rel in $jobs) {
  $docx = Join-Path $sets $rel
  $pdf = [System.IO.Path]::ChangeExtension($docx, ".pdf")
  Write-Output "TRY $rel"
  $done = $false
  for ($a = 1; $a -le 3 -and -not $done; $a++) {
    Kill-Word
    $word = $null
    $doc = $null
    try {
      $word = New-Object -ComObject Word.Application
      $word.Visible = $false
      $word.DisplayAlerts = 0
      Start-Sleep -Seconds 1
      if (Test-Path $pdf) { Remove-Item $pdf -Force -ErrorAction SilentlyContinue }
      $doc = $word.Documents.Open($docx, $false, $true)
      Start-Sleep -Milliseconds 800
      # ExportAsFixedFormat: 17 = PDF
      $doc.ExportAsFixedFormat($pdf, 17, $false, 0, $false, 0, 0, $false, $true, $false, 0, $true, $true, $false)
      Start-Sleep -Milliseconds 500
      if ((Test-Path $pdf) -and ((Get-Item $pdf).Length -gt 0)) {
        Write-Output "  attempt $a OK size=$((Get-Item $pdf).Length)"
        $done = $true
        $ok++
      } else {
        Write-Output "  attempt $a EMPTY"
      }
    } catch {
      Write-Output "  attempt $a ERR: $($_.Exception.Message)"
    } finally {
      try { if ($doc) { $doc.Close([ref]$false) | Out-Null } } catch {}
      try { if ($word) { $word.Quit([ref]$false) | Out-Null } } catch {}
      try {
        if ($doc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
        if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
      } catch {}
      [GC]::Collect(); [GC]::WaitForPendingFinalizers()
      Kill-Word
    }
  }
  if (-not $done) { $fail++ }
}
Write-Output "FINAL ok=$ok fail=$fail"
