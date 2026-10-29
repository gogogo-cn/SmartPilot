<#
  build-web-setup.ps1 - Build the desktop edition self-extracting installer
  (SmartPilotWebSetup.exe) from dist\SmartPilotWeb-Portable.zip.
  Prerequisite: run Tools\package-web.ps1 first.
#>
param(
    [string]$ZipFile = '',
    [string]$OutputDir = ''
)
$ErrorActionPreference = 'Stop'
if (-not $ZipFile)   { $ZipFile   = Join-Path $PSScriptRoot '..\dist\SmartPilotWeb-Portable.zip' }
if (-not $OutputDir) { $OutputDir = Join-Path $PSScriptRoot '..\dist' }

$cs = Join-Path $PSScriptRoot 'selfext\WebExtractor.cs'
if (-not (Test-Path $ZipFile)) { throw "ZIP not found: $ZipFile. Run package-web.ps1 first." }

$csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) { $csc = 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe' }

$out = Join-Path $OutputDir 'SmartPilotWebSetup.exe'
$resource = '/resource:' + $ZipFile + ',selfext.zip'

Write-Host ('Compiling desktop edition installer (' + [math]::Round((Get-Item $ZipFile).Length/1MB,1) + ' MB embedded)...') -ForegroundColor Cyan
& $csc /target:winexe /out:$out /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll $resource $cs 2>&1 | Out-Host
Write-Host ('CSC_EXIT=' + $LASTEXITCODE)

if (Test-Path $out) {
    Write-Host ('SETUP_EXE -> ' + $out + ' (' + [math]::Round((Get-Item $out).Length/1MB,1) + ' MB)') -ForegroundColor Green
} else {
    Write-Host 'FAILED: SmartPilotWebSetup.exe not produced.' -ForegroundColor Red
}
