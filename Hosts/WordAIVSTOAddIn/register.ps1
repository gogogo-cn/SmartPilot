# ============================================================================
# Word VSTO add-in registration (HKCU only, no admin required)
#
# Registers the add-in with vstolocal so it loads directly from
# bin\Debug\WordAIVSTOAddIn.vsto (no ClickOnce install needed).
#
# Usage:
#   .\register.ps1               register + enable
#   .\register.ps1 -Unregister   unregister
# After registration, fully exit Word and start it again.
# ============================================================================
param([switch]$Unregister)

$AddInName = 'WordAIVSTOAddIn'
$Manifest = (Join-Path $PSScriptRoot 'bin\Debug\WordAIVSTOAddIn.vsto')
$Key = "HKCU:\Software\Microsoft\Office\Word\Addins\$AddInName"

if ($Unregister) {
    if (Test-Path $Key) { Remove-Item $Key -Recurse -Force }
    Write-Host "Unregistered. Please restart Word." -ForegroundColor Green
    exit 0
}

if (-not (Test-Path $Manifest)) {
    Write-Host "NOT FOUND: $Manifest - build the project (Debug) first." -ForegroundColor Red
    exit 1
}

$manifestUri = 'file:///' + ($Manifest -replace '\\', '/') + '|vstolocal'

New-Item -Path $Key -Force | Out-Null
New-ItemProperty -Path $Key -Name 'Description' -Value 'SmartPilot' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $Key -Name 'FriendlyName' -Value 'SmartPilot' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $Key -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $Key -Name 'Manifest' -Value $manifestUri -PropertyType String -Force | Out-Null

Write-Host ""
Write-Host "Registered: $AddInName" -ForegroundColor Green
Write-Host "  manifest: $manifestUri" -ForegroundColor Gray
Write-Host "Fully exit WINWORD.EXE, then start Word again." -ForegroundColor Green
