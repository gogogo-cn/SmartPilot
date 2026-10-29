# ============================================================================
# PowerPoint COM add-in registration (no admin required)
#
#   1) COM class registration: try "RegAsm /codebase" (needs admin);
#      on failure fall back to "RegAsm /regfile" -> rewrite HKCR to
#      HKCU\Software\Classes -> reg import (per-user COM registration)
#   2) Office add-in key: HKCU\Software\Microsoft\Office\PowerPoint\Addins\<ProgId>
#
# Usage:
#   .\register.ps1               register + enable
#   .\register.ps1 -Unregister   unregister
# After registration, fully exit PowerPoint and start it again.
# ============================================================================
param([switch]$Unregister)

$ErrorActionPreference = 'Stop'
$Dll = Join-Path $PSScriptRoot 'bin\Debug\PptAIVSTOAddIn.dll'
$ProgId = 'PptAIVSTOAddIn.Connect'
$AddinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$ProgId"
$ClsRoot = "HKCU:\Software\Classes\$ProgId"

$regasm64 = Join-Path $env:WinDir 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
$regasm32 = Join-Path $env:WinDir 'Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'
$RegAsm = if (Test-Path $regasm64) { $regasm64 } else { $regasm32 }

if (-not (Test-Path $Dll)) {
    Write-Host "NOT FOUND: $Dll - build the project (Debug) first." -ForegroundColor Red
    exit 1
}

# ---------------- unregister ----------------
if ($Unregister) {
    Write-Host "Unregistering COM add-in..." -ForegroundColor Yellow
    try { & $RegAsm /unregister $Dll | Out-Null } catch { Write-Host "RegAsm unregister skipped." -ForegroundColor DarkYellow }
    if (Test-Path $AddinKey) { Remove-Item $AddinKey -Recurse -Force }
    if (Test-Path $ClsRoot) { Remove-Item $ClsRoot -Recurse -Force }
    Write-Host "Done. Please restart PowerPoint." -ForegroundColor Green
    exit 0
}

# ---------------- 1) COM class ----------------
Write-Host "RegAsm: $RegAsm" -ForegroundColor Gray
$comRegistered = $false

& $RegAsm /codebase $Dll 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
if ($LASTEXITCODE -eq 0) {
    $comRegistered = $true
    Write-Host "COM class registered (HKCR)." -ForegroundColor Green
} else {
    Write-Host "Direct registration failed (likely permission). Falling back to per-user HKCU..." -ForegroundColor Yellow
    $regFile = Join-Path $env:TEMP "$ProgId.reg"
    if (Test-Path $regFile) { Remove-Item $regFile -Force }
    & $RegAsm /codebase /regfile:$regFile $Dll 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
    if (-not (Test-Path $regFile)) {
        Write-Host "RegAsm could not produce a .reg file." -ForegroundColor Red
        exit 1
    }
    $content = [System.IO.File]::ReadAllText($regFile)
    $content = $content -replace 'HKEY_CLASSES_ROOT', 'HKEY_CURRENT_USER\Software\Classes'
    $outFile = Join-Path $env:TEMP "$ProgId.hkcu.reg"
    [System.IO.File]::WriteAllText($outFile, $content, [System.Text.Encoding]::Unicode)
    & reg import $outFile 2>&1 | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Per-user COM import failed." -ForegroundColor Red
        exit $LASTEXITCODE
    }
    $comRegistered = $true
    Write-Host "COM class registered (HKCU per-user)." -ForegroundColor Green
}

# ---------------- 2) Office add-in key ----------------
New-Item -Path $AddinKey -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name 'Description' -Value 'SmartPilot' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name 'FriendlyName' -Value 'SmartPilot' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $AddinKey -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null

Write-Host ""
Write-Host "Registered: $ProgId" -ForegroundColor Green
Write-Host "  add-in key : $AddinKey" -ForegroundColor Gray
Write-Host "  COM class  : $comRegistered" -ForegroundColor Gray
Write-Host "Fully exit POWERPNT.EXE, then start PowerPoint again." -ForegroundColor Green
