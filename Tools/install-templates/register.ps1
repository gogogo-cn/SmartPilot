<#
  ExcelAIVSTOAddIn unified per-user installer (HKCU, no administrator required).

  Registers Excel + Word (VSTO add-ins) and PowerPoint (COM add-in) so Office
  loads them directly from this folder via the |vstolocal offline mechanism.
  Also imports the self-signed code-signing certificate into
  CurrentUser\TrustedPublisher so Office trusts the manifests.

  Usage:
    powershell -ExecutionPolicy Bypass -File register.ps1
    powershell -ExecutionPolicy Bypass -File register.ps1 -Unregister
#>
param([switch]$Unregister)

$ErrorActionPreference = 'Stop'
$scriptDir = $PSScriptRoot

function Write-Step($m, $c = 'Cyan') { Write-Host $m -ForegroundColor $c }

function Ensure-CertTrusted {
    $pfx = Get-ChildItem -Path (Join-Path $scriptDir 'certs') -Filter *.pfx -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $pfx) { Write-Step 'No .pfx found in certs/, skipping cert import (manifests may be untrusted).' Yellow; return }
    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($pfx.FullName, '', 'DefaultKeySet')
    } catch {
        Write-Step ("Could not load certificate '$($pfx.Name)': $_") Red; return
    }
    foreach ($storeName in @('TrustedPublisher', 'My')) {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, 'CurrentUser')
        $store.Open('ReadWrite')
        if (-not ($store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint })) {
            $store.Add($cert)
            Write-Step "Imported cert to CurrentUser\$storeName" Green
        } else {
            Write-Step "Cert already in CurrentUser\$storeName" Gray
        }
        $store.Close()
    }
}

function Register-VstoAddin($app, $name, $vstoRelPath) {
    $vsto = Join-Path $scriptDir $vstoRelPath
    if (-not (Test-Path $vsto)) { Write-Step "MISSING: $vsto" Red; return $false }
    $uri = 'file:///' + ($vsto -replace '\\', '/') + '|vstolocal'
    $key = "HKCU:\Software\Microsoft\Office\$app\Addins\$name"
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name Manifest -Value $uri -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name FriendlyName -Value 'SmartPilot' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name Description -Value 'SmartPilot AI Assistant' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
    Write-Step "Registered $app\$name -> $uri" Green
    return $true
}

function Unregister-VstoAddin($app, $name) {
    $key = "HKCU:\Software\Microsoft\Office\$app\Addins\$name"
    if (Test-Path $key) { Remove-Item $key -Recurse -Force; Write-Step "Unregistered $app\$name" Green }
}

function Register-PptCom {
    $dll = Join-Path $scriptDir 'PowerPoint\PptAIVSTOAddIn.dll'
    if (-not (Test-Path $dll)) { Write-Step "MISSING: $dll" Red; return $false }
    $progId = 'PptAIVSTOAddIn.Connect'
    $addinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$progId"

    $regasm64 = Join-Path $env:WinDir 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
    $regasm32 = Join-Path $env:WinDir 'Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'
    $regasm = if (Test-Path $regasm64) { $regasm64 } else { $regasm32 }
    if (-not (Test-Path $regasm)) { Write-Step "RegAsm.exe not found; cannot register PowerPoint COM add-in." Red; return $false }

    # Export to a .reg (with CodeBase) without writing to HKCR, then import into HKCU (per-user).
    $regFile = Join-Path $env:TEMP ($progId + '.reg')
    if (Test-Path $regFile) { Remove-Item $regFile -Force }
    # reg.exe/RegAsm write their status lines to stderr; with EAP=Stop + 2>&1 that
    # becomes a terminating NativeCommandError in PS 5.1, so relax EAP locally.
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    & $regasm /codebase /regfile:$regFile $dll 2>&1 | ForEach-Object { Write-Step "  regasm: $_" Gray }
    $ErrorActionPreference = $prevEap
    if (-not (Test-Path $regFile)) { Write-Step 'RegAsm failed to produce a .reg file.' Red; return $false }
    $content = [System.IO.File]::ReadAllText($regFile)
    $content = $content -replace 'HKEY_CLASSES_ROOT', 'HKEY_CURRENT_USER\Software\Classes'
    $outFile = Join-Path $env:TEMP ($progId + '.hkcu.reg')
    [System.IO.File]::WriteAllText($outFile, $content, [System.Text.Encoding]::Unicode)
    $regExe = Join-Path $env:WinDir 'System32\reg.exe'
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    & $regExe import $outFile 2>&1 | ForEach-Object { Write-Step "  import: $_" Gray }
    $ErrorActionPreference = $prevEap
    if ($LASTEXITCODE -ne 0) { Write-Step 'Per-user COM import failed.' Red; return $false }

    New-Item -Path $addinKey -Force | Out-Null
    New-ItemProperty -Path $addinKey -Name Description -Value 'SmartPilot' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $addinKey -Name FriendlyName -Value 'SmartPilot' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $addinKey -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
    Write-Step 'Registered PowerPoint COM add-in (per-user).' Green
    return $true
}

function Unregister-PptCom {
    $progId = 'PptAIVSTOAddIn.Connect'
    $addinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$progId"
    $clsRoot = "HKCU:\Software\Classes\$progId"
    if (Test-Path $addinKey) { Remove-Item $addinKey -Recurse -Force }
    if (Test-Path $clsRoot) { Remove-Item $clsRoot -Recurse -Force }
    Write-Step 'Unregistered PowerPoint COM add-in.' Green
}

function Clear-Resiliency {
    foreach ($ver in @('16.0', '15.0')) {
        foreach ($app in @('Excel', 'Word', 'PowerPoint')) {
            $rp = "HKCU:\Software\Microsoft\Office\$ver\$app\Resiliency\DisabledItems"
            if (Test-Path $rp) { Remove-Item $rp -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}

function Check-Prereqs {
    Write-Step '--- System prerequisites ---' Cyan
    $vsto = (Test-Path 'HKLM:\SOFTWARE\Microsoft\VSTO Runtime Setup') -or (Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VSTO Runtime Setup')
    if ($vsto) { Write-Step 'VSTO Runtime: present' Green }
    else { Write-Step 'VSTO Runtime: MISSING - install "Microsoft Visual Studio 2010 Tools for Office Runtime" (needs admin once): https://www.microsoft.com/en-us/download/details.aspx?id=56961' Yellow }

    $webv = Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}' -ErrorAction SilentlyContinue
    if ($webv) { Write-Step ('WebView2 Runtime: present (' + $webv.pv + ')') Green }
    else { Write-Step 'WebView2 Runtime: MISSING - install "Microsoft Edge WebView2 Runtime" (Evergreen): https://go.microsoft.com/fwlink/p/?LinkId=2124703' Yellow }

    $net = Test-Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full'
    if ($net) { Write-Step '.NET Framework 4.7.2+: present' Green }
    else { Write-Step '.NET Framework 4.7.2+: MISSING - install from https://dotnet.microsoft.com/download/dotnet-framework' Yellow }
}

if ($Unregister) {
    Unregister-VstoAddin 'Excel' 'ExcelAIVSTOAddIn'
    Unregister-VstoAddin 'Word' 'WordAIVSTOAddIn'
    Unregister-PptCom
    $uninsKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\SmartPilot for Office'
    if (Test-Path $uninsKey) { Remove-Item $uninsKey -Recurse -Force }
    Write-Step 'All add-ins unregistered. Please restart Office.' Green
    exit 0
}

Write-Step '=== SmartPilot for Office installer (per-user) ===' Cyan
Ensure-CertTrusted
$fail = 0
if (-not (Register-VstoAddin 'Excel' 'ExcelAIVSTOAddIn' 'Excel\ExcelAIVSTOAddIn.vsto')) { $fail++ }
if (-not (Register-VstoAddin 'Word' 'WordAIVSTOAddIn' 'Word\WordAIVSTOAddIn.vsto')) { $fail++ }
if (-not (Register-PptCom)) { $fail++ }
Clear-Resiliency
Check-Prereqs
Write-Step ''
if ($fail -gt 0) {
    Write-Step ("Installation finished with " + $fail + " failed component(s). See red messages above.") Red
    exit 1
}

# Standard uninstall entry (Settings -> Apps / Control Panel), per-user, no admin.
$uninsKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\SmartPilot for Office'
New-Item -Path $uninsKey -Force | Out-Null
New-ItemProperty -Path $uninsKey -Name DisplayName -Value 'SmartPilot for Office' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninsKey -Name DisplayVersion -Value '1.0.0' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninsKey -Name Publisher -Value 'SmartPilot' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninsKey -Name InstallLocation -Value $scriptDir -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninsKey -Name UninstallString -Value (Join-Path $scriptDir 'uninstall.bat') -PropertyType String -Force | Out-Null
New-ItemProperty -Path $uninsKey -Name NoModify -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $uninsKey -Name NoRepair -Value 1 -PropertyType DWord -Force | Out-Null
Write-Step 'Registered uninstall entry (Settings -> Apps -> SmartPilot for Office).' Green

Write-Step 'Installation complete. Fully exit all Office apps, then reopen Excel/Word/PowerPoint.' Green
