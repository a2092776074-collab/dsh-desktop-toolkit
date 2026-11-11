# 重新编译 C# 工具（需要 .NET Framework 4 的 csc.exe，Windows 自带）
# 用法：powershell -ExecutionPolicy Bypass -File build.ps1
$ErrorActionPreference = 'Stop'
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) { $csc = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe" }
if (-not (Test-Path $csc)) { throw "找不到 csc.exe" }

$gac = "C:\Windows\Microsoft.NET\assembly\GAC_MSIL"
$uiaRefs = @(
  "/r:$gac\UIAutomationClient\v4.0_4.0.0.0__31bf3856ad364e35\UIAutomationClient.dll",
  "/r:$gac\UIAutomationTypes\v4.0_4.0.0.0__31bf3856ad364e35\UIAutomationTypes.dll",
  "/r:$gac\WindowsBase\v4.0_4.0.0.0__31bf3856ad364e35\WindowsBase.dll"
)

New-Item -ItemType Directory -Force -Path bin | Out-Null

& $csc /nologo /target:exe /out:bin\mouse.exe   src\MouseTool.cs
& $csc /nologo /target:exe /out:bin\grid.exe    src\GridTool.cs /r:System.Windows.Forms.dll /r:System.Drawing.dll
& $csc /nologo /target:exe /out:bin\imgdiff.exe src\ImgDiff.cs /r:System.Drawing.dll
& $csc /nologo /target:exe /out:bin\window.exe  src\WindowTool.cs
& $csc /nologo /target:exe /out:bin\uia.exe     @uiaRefs src\UiaTool.cs

Get-ChildItem bin\*.exe | Select-Object Name, Length
Write-Output "DONE"
