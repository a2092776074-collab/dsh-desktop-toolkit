param([int]$Port = 41983, [int]$IdleMinutes = 10)
$ErrorActionPreference = 'Stop'

# 2026-11-11: resident OCR service, auto exit after idle minutes
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

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if (-not $engine) { Write-Output '{"ok":false,"error":"no OCR engine"}'; exit 1 }

$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
$listener.Start()
Write-Output ("READY " + $Port)

function Handle-Request([string]$line) {
  $parts = $line -split "`t"
  $img = $parts[0]
  $find = if ($parts.Count -gt 1) { $parts[1] } else { '' }
  $scale = if ($parts.Count -gt 2) { [double]$parts[2] } else { 1 }
  if ($line -eq 'quit') { return '{"ok":true,"bye":true}' }
  try {
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync((Resolve-Path $img).Path)) ([Windows.Storage.StorageFile])
    $st = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($st)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bmp = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $res = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
    $div = if ($scale -gt 0) { $scale } else { 1 }
    $lines = @()
    foreach ($ln in $res.Lines) {
      $minX = 1e9; $minY = 1e9; $maxX = -1e9; $maxY = -1e9
      foreach ($w in $ln.Words) {
        $r = $w.BoundingRect
        if ($r.X -lt $minX) { $minX = $r.X }
        if ($r.Y -lt $minY) { $minY = $r.Y }
        if (($r.X + $r.Width) -gt $maxX) { $maxX = $r.X + $r.Width }
        if (($r.Y + $r.Height) -gt $maxY) { $maxY = $r.Y + $r.Height }
      }
      $lines += [pscustomobject]@{ text = $ln.Text; x = [int]($minX / $div); y = [int]($minY / $div); w = [int](($maxX - $minX) / $div); h = [int](($maxY - $minY) / $div) }
    }
    if ($find) {
      $target = $find -replace '\s', ''
      $hits = $lines | Where-Object { ($_.text -replace '\s', '') -like "*$target*" }
      $hit = $hits | Sort-Object { [math]::Abs((($_.text -replace '\s','').Length) - $target.Length) } | Select-Object -First 1
      if ($hit) {
        return ([pscustomobject]@{ ok = $true; text = $hit.text; x = $hit.x; y = $hit.y; cx = [int]($hit.x + $hit.w / 2); cy = [int]($hit.y + $hit.h / 2) } | ConvertTo-Json -Compress)
      }
      return '{"ok":false,"error":"text not found"}'
    }
    return ([pscustomobject]@{ ok = $true; count = $lines.Count; lines = $lines } | ConvertTo-Json -Compress -Depth 5)
  } catch {
    return ('{"ok":false,"error":' + ($_.Exception.Message | ConvertTo-Json) + '}')
  }
}

$last = Get-Date
while ($true) {
  if ($listener.Pending()) {
    $last = Get-Date
    $client = $listener.AcceptTcpClient()
    try {
      $stream = $client.GetStream()
      $reader = New-Object System.IO.StreamReader($stream)
      $writer = New-Object System.IO.StreamWriter($stream)
      $writer.AutoFlush = $true
      while ($true) {
        $line = $reader.ReadLine()
        if ($null -eq $line) { break }
        $line = $line.Trim()
        if ($line -eq '') { continue }
        $resp = Handle-Request $line
        $writer.WriteLine($resp)
        if ($line -eq 'quit') { break }
      }
    } catch {
    } finally {
      $client.Close()
    }
  } else {
    Start-Sleep -Milliseconds 500
    if (((Get-Date) - $last).TotalMinutes -ge $IdleMinutes) {
      Write-Output 'IDLE EXIT'
      break
    }
  }
}
$listener.Stop()
