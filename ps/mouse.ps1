param(
  [Parameter(Mandatory=$true)][ValidateSet('click','drag','move')][string]$Action,
  [int]$X1 = 0, [int]$Y1 = 0, [int]$X2 = 0, [int]$Y2 = 0,
  [int]$Steps = 20, [int]$DelayMs = 15, [string]$ShotDir = '', [switch]$DryRun, [switch]$MidShot
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class MInput {
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e);
  public const uint LEFTDOWN = 0x0002, LEFTUP = 0x0004, MOVE = 0x0001;
  public static void Down() { mouse_event(LEFTDOWN, 0, 0, 0, UIntPtr.Zero); }
  public static void Up() { mouse_event(LEFTUP, 0, 0, 0, UIntPtr.Zero); }
  public static void To(int x, int y) { SetCursorPos(x, y); System.Threading.Thread.Sleep(5); }
}
"@

function Shot([string]$path) {
  $b = [System.Windows.Forms.SystemInformation]::VirtualScreen
  $bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($b.Left, $b.Top, 0, 0, $bmp.Size)
  $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose()
}

if ($DryRun) {
  Write-Output ("DRY {0} ({1},{2})->({3},{4}) steps={5} delay={6}ms" -f $Action, $X1, $Y1, $X2, $Y2, $Steps, $DelayMs)
  exit 0
}

$before = $null; $after = $null
if ($ShotDir) {
  New-Item -ItemType Directory -Force $ShotDir | Out-Null
  $before = Join-Path $ShotDir ("before-" + (Get-Date -Format 'HHmmssfff') + ".png")
  Shot $before
}

switch ($Action) {
  'move' {
    [MInput]::To($X1, $Y1)
  }
  'click' {
    [MInput]::To($X1, $Y1)
    Start-Sleep -Milliseconds 80
    [MInput]::Down(); Start-Sleep -Milliseconds 60; [MInput]::Up()
  }
  'drag' {
    [MInput]::To($X1, $Y1)
    Start-Sleep -Milliseconds 120
    [MInput]::Down()
    Start-Sleep -Milliseconds 120
    for ($i = 1; $i -le $Steps; $i++) {
      $x = [int]($X1 + ($X2 - $X1) * $i / $Steps)
      $y = [int]($Y1 + ($Y2 - $Y1) * $i / $Steps)
      [MInput]::To($x, $y)
      Start-Sleep -Milliseconds $DelayMs
      if ($MidShot -and $ShotDir -and $i -eq [int]($Steps * 0.8)) {
        $mid = Join-Path $ShotDir ("mid-" + (Get-Date -Format 'HHmmssfff') + ".png")
        Shot $mid
        Write-Output ("MID " + (Split-Path $mid -Leaf))
      }
    }
    Start-Sleep -Milliseconds 150
    [MInput]::Up()
  }
}

if ($ShotDir) {
  Start-Sleep -Milliseconds 250
  $after = Join-Path $ShotDir ("after-" + (Get-Date -Format 'HHmmssfff') + ".png")
  Shot $after
  $h1 = (Get-FileHash $before -Algorithm MD5).Hash
  $h2 = (Get-FileHash $after -Algorithm MD5).Hash
  Write-Output ("OK {0} before={1} after={2} changed={3}" -f $Action, (Split-Path $before -Leaf), (Split-Path $after -Leaf), ($h1 -ne $h2))
} else {
  Write-Output ("OK {0}" -f $Action)
}
