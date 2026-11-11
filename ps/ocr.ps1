param(
  [Parameter(Mandatory=$true)][string]$Image,
  [string]$Find = '',
  [double]$Scale = 1
)
$ErrorActionPreference = 'Stop'
try {
  [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
  [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime] | Out-Null
  [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType=WindowsRuntime] | Out-Null
  Add-Type -AssemblyName System.Runtime.WindowsRuntime | Out-Null

  $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
      $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    })[0]

  function Await($op, $type) {
    $m = $asTaskGeneric.MakeGenericMethod($type)
    $t = $m.Invoke($null, @($op))
    $t.Wait(-1) | Out-Null
    $t.Result
  }

  $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync((Resolve-Path $Image).Path)) ([Windows.Storage.StorageFile])
  $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
  $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $bmp = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

  $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
  if (-not $engine) { Write-Output '{"ok":false,"error":"no OCR engine for current language"}'; exit 1 }

  $res = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
  $div = if ($Scale -gt 0) { $Scale } else { 1 }
  $lines = @()
  foreach ($line in $res.Lines) {
    $minX = 1e9; $minY = 1e9; $maxX = -1e9; $maxY = -1e9
    foreach ($w in $line.Words) {
      $r = $w.BoundingRect
      if ($r.X -lt $minX) { $minX = $r.X }
      if ($r.Y -lt $minY) { $minY = $r.Y }
      if (($r.X + $r.Width) -gt $maxX) { $maxX = $r.X + $r.Width }
      if (($r.Y + $r.Height) -gt $maxY) { $maxY = $r.Y + $r.Height }
    }
    $lines += [pscustomobject]@{
      text = $line.Text
      x = [int]($minX / $div); y = [int]($minY / $div)
      w = [int](($maxX - $minX) / $div); h = [int](($maxY - $minY) / $div)
    }
  }

  if ($Find) {
    $target = $Find -replace '\s', ''
    $hits = $lines | Where-Object { ($_.text -replace '\s', '') -like "*$target*" }
    $hit = $hits | Sort-Object { [math]::Abs((($_.text -replace '\s','').Length) - $target.Length) } | Select-Object -First 1
    if (-not $hit) { Write-Output ('{"ok":false,"error":"text not found","count":' + $lines.Count + '}'); exit 1 }
    $cx = [int]($hit.x + $hit.w / 2); $cy = [int]($hit.y + $hit.h / 2)
    Write-Output ([pscustomobject]@{ ok = $true; text = $hit.text; x = $hit.x; y = $hit.y; cx = $cx; cy = $cy } | ConvertTo-Json -Compress)
  } else {
    Write-Output ([pscustomobject]@{ ok = $true; count = $lines.Count; lines = $lines } | ConvertTo-Json -Compress -Depth 5)
  }
} catch {
  Write-Output ('{"ok":false,"error":' + ($_.Exception.Message | ConvertTo-Json) + '}')
  exit 1
}
