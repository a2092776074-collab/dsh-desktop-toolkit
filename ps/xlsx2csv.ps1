param(
  [Parameter(Mandatory=$true)][string]$In,
  [Parameter(Mandatory=$true)][string]$Out
)
$ErrorActionPreference = 'Stop'
$excel = $null
$wb = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open((Resolve-Path $In).Path, 0, $true)
  $wb.SaveAs([System.IO.Path]::GetFullPath($Out), 6)  # 6 = xlCSV
  $wb.Close($false)
  $excel.Quit()
  Write-Output "OK"
} catch {
  if ($wb) { try { $wb.Close($false) } catch {} }
  if ($excel) { try { $excel.Quit() } catch {} }
  Write-Output ("ERR " + $_.Exception.Message)
  exit 1
} finally {
  if ($wb) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
  if ($excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
