<#
  build-iexpress-setup.ps1 - Build a single-click self-extracting installer
  (SmartPilotSetup.exe) using Windows' built-in IExpress.

  The installer extracts the staged package to
  %LOCALAPPDATA%\Programs\ExcelAIVSTOAddIn and then auto-runs install.bat
  (which registers the three Office add-ins into HKCU and imports the cert).
  No administrator required; one double-click installs everything.

  Prerequisite: run Tools\package-addin.ps1 first (produces dist\ExcelAIVSTOAddIn).
#>
param(
    [string]$StageDir  = '',
    [string]$OutputDir = ''
)
$ErrorActionPreference = 'Stop'
if (-not $StageDir)  { $StageDir  = Join-Path $PSScriptRoot '..\dist\ExcelAIVSTOAddIn' }
if (-not $OutputDir) { $OutputDir = Join-Path $PSScriptRoot '..\dist' }

$installDir = '%LOCALAPPDATA%\Programs\ExcelAIVSTOAddIn'
$sed = Join-Path $OutputDir 'SmartPilotSetup.sed'

if (-not (Test-Path $StageDir)) { throw "Stage dir not found: $StageDir. Run package-addin.ps1 first." }

# Enumerate every file (relative path, backslash separated) under the stage dir.
$files = Get-ChildItem -Path $StageDir -Recurse -File | ForEach-Object {
    $_.FullName.Substring($StageDir.Length).TrimStart('\', '/').Replace('/', '\')
}
if ($files.Count -eq 0) { throw "No files found under $StageDir" }

$head = @"
[Version]
Signature=`"`$Windows NT`$"
Class=IEXPRESS
Provider="SmartPilot"
[Strings]
ProductName="SmartPilot for Office"
PackageName="SmartPilotSetup.exe"
DefaultDir="$installDir"
AppLaunched="cmd /c ""$installDir\install.bat"""
PostInstallCmd="<None>"
Title="SmartPilot for Office 安装"
[Options]
HideExtractAnimation=1
HideFinalWindow=1
[Files]
SourceFiles=$StageDir
[$StageDir]
"@
$content = $head + ($files -join "`r`n") + "`r`n"
[System.IO.File]::WriteAllText($sed, $content, [System.Text.Encoding]::Default)
Write-Host ("SED written: $sed (" + $files.Count + " files)") -ForegroundColor Green

$iex = 'C:\Windows\System32\iexpress.exe'
if (-not (Test-Path $iex)) { throw "IExpress not found at $iex" }
Write-Host 'Building self-extracting installer (this may take a minute for ~105 MB)...' -ForegroundColor Cyan
& $iex /N /Q $sed 2>&1 | Out-Host
Write-Host ('IEXPRESS_EXIT=' + $LASTEXITCODE)

$exe = Join-Path $OutputDir 'SmartPilotSetup.exe'
if (Test-Path $exe) {
    Write-Host ('SETUP_EXE -> ' + $exe + ' (' + [math]::Round((Get-Item $exe).Length / 1MB, 1) + ' MB)') -ForegroundColor Green
} else {
    Write-Host 'FAILED: SmartPilotSetup.exe was not produced.' -ForegroundColor Red
}
