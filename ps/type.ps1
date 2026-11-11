param(
  [Parameter(Mandatory=$true)][string]$Text,
  [switch]$NoPaste
)
$ErrorActionPreference = 'Stop'
try {
  # 中文输入走剪贴板粘贴（SendKeys 直接发中文会丢）
  Set-Clipboard -Value $Text
  Start-Sleep -Milliseconds 120
  $ws = New-Object -ComObject WScript.Shell
  if (-not $NoPaste) { $ws.SendKeys('^v') }
  Write-Output ('OK ' + $Text.Length + ' chars')
} catch {
  Write-Output ('ERR ' + $_.Exception.Message)
  exit 1
}
