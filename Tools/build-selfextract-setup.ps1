<#
  build-selfextract-setup.ps1 - Build a single-click self-extracting installer
  (SmartPilotSetup.exe) using the .NET Framework C# compiler (csc.exe).

  The existing portable ZIP (dist\ExcelAIVSTOAddIn-Portable.zip) is embedded as
  a manifest resource into a tiny C# bootstrapper. Double-clicking the EXE
  extracts to %LOCALAPPDATA%\Programs\ExcelAIVSTOAddIn and runs install.bat
  (registers Excel/Word/PPT add-ins into HKCU, imports the cert). No admin.

  Prerequisite: run Tools\package-addin.ps1 first (produces the ZIP).
#>
param(
    [string]$ZipFile = '',
    [string]$OutputDir = ''
)
$ErrorActionPreference = 'Stop'
if (-not $ZipFile)   { $ZipFile   = Join-Path $PSScriptRoot '..\dist\ExcelAIVSTOAddIn-Portable.zip' }
if (-not $OutputDir) { $OutputDir = Join-Path $PSScriptRoot '..\dist' }

$cs = Join-Path $PSScriptRoot 'selfext\SelfExtractor.cs'
if (-not (Test-Path $ZipFile)) { throw "ZIP not found: $ZipFile. Run package-addin.ps1 first." }
if (-not (Test-Path $cs))      { throw "SelfExtractor.cs not found: $cs" }

$csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) { $csc = 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe' }
if (-not (Test-Path $csc)) { throw "csc.exe not found (need .NET Framework 4.x)" }

$out = Join-Path $OutputDir 'SmartPilotSetup.exe'
$resource = '/resource:' + $ZipFile + ',selfext.zip'

Write-Host ('Compiling self-extracting installer (' + [math]::Round((Get-Item $ZipFile).Length/1MB,1) + ' MB embedded)...') -ForegroundColor Cyan
& $csc /target:winexe /out:$out /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll $resource $cs 2>&1 | Out-Host
Write-Host ('CSC_EXIT=' + $LASTEXITCODE)

if (Test-Path $out) {
    Write-Host ('SETUP_EXE -> ' + $out + ' (' + [math]::Round((Get-Item $out).Length/1MB,1) + ' MB)') -ForegroundColor Green
} else {
    Write-Host 'FAILED: SmartPilotSetup.exe not produced.' -ForegroundColor Red
}
