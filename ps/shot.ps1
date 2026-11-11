param(
  [Parameter(Mandatory=$true)][string]$Out,
  [int]$X = -2147483648, [int]$Y = -2147483648, [int]$W = 0, [int]$H = 0
)
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
$vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
$rx = $X; $ry = $Y; $rw = $W; $rh = $H
if ($X -eq -2147483648) { $rx = $vs.Left; $ry = $vs.Top; $rw = $vs.Width; $rh = $vs.Height }
if ($rw -le 0 -or $rh -le 0) { $rw = $vs.Width; $rh = $vs.Height; $rx = $vs.Left; $ry = $vs.Top }
$bmp = New-Object System.Drawing.Bitmap $rw, $rh
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($rx, $ry, 0, 0, $bmp.Size)
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose()
$bmp.Dispose()
Write-Output ("OK " + $Out + " " + $rx + "," + $ry + " " + $rw + "x" + $rh)
