param(
  [Parameter(Mandatory=$true)][string]$In,
  [Parameter(Mandatory=$true)][string]$Out
)
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open((Resolve-Path $In).Path, $false, $true)
  $doc.SaveAs([ref]([System.IO.Path]::GetFullPath($Out)), [ref]17)  # 17 = wdFormatPDF
  $doc.Close($false)
  $word.Quit()
  Write-Output "OK"
} catch {
  if ($doc) { try { $doc.Close($false) } catch {} }
  if ($word) { try { $word.Quit() } catch {} }
  Write-Output ("ERR " + $_.Exception.Message)
  exit 1
} finally {
  if ($doc) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }
  if ($word) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
}
