<#
  package-addin.ps1 - Build a distributable package for ExcelAIVSTOAddIn.
  Produces BOTH:
    * a portable ZIP  (ExcelAIVSTOAddIn-Portable.zip)  - extract anywhere + run install.bat
    * an Inno Setup script (ExcelAIVSTOAddIn.iss); if iscc.exe is available it also
      compiles ExcelAIVSTOAddIn-setup.exe (per-user install under %LocalAppData%).

  Per-user (HKCU) registration, no administrator required. Uses the self-signed
  dev certificate shipped in certs\ (imported to TrustedPublisher at install time).

  Usage:
    Tools\package-addin.ps1                      # default: Debug artifacts, both outputs
    Tools\package-addin.ps1 -SourceConfig Release
    Tools\package-addin.ps1 -NoInno              # zip only
    Tools\package-addin.ps1 -NoZip               # inno only
#>
param(
    [string]$SourceConfig = 'Debug',
    [string]$OutputDir    = '',
    [string]$CertPfx      = '',
    [switch]$NoZip,
    [switch]$NoInno
)

$ErrorActionPreference = 'Stop'
$scriptRoot = $PSScriptRoot
$repoRoot   = Resolve-Path (Join-Path $scriptRoot '..')

if (-not $OutputDir) { $OutputDir = Join-Path $repoRoot 'dist' }
if (-not $CertPfx)   { $CertPfx   = Join-Path $repoRoot 'ExcelAIVSTOAddIn_TemporaryKey.pfx' }

$out   = Resolve-Path $OutputDir -ErrorAction SilentlyContinue
if (-not $out) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null; $out = Resolve-Path $OutputDir }
$stage = Join-Path $out 'ExcelAIVSTOAddIn'

Write-Host "Repo root : $repoRoot" -ForegroundColor Cyan
Write-Host "Output    : $out" -ForegroundColor Cyan
Write-Host "Source cfg: $SourceConfig" -ForegroundColor Cyan
Write-Host "Stage     : $stage" -ForegroundColor Cyan

# ---- 1. prepare staging dir ----
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Path $stage -Force | Out-Null

# ---- 2. collect host runtimes ----
$sources = @{
    'Excel'      = (Join-Path $repoRoot "bin\$SourceConfig")
    'Word'       = (Join-Path $repoRoot "Hosts\WordAIVSTOAddIn\bin\$SourceConfig")
    'PowerPoint' = (Join-Path $repoRoot "Hosts\PptAIVSTOAddIn\bin\$SourceConfig")
}
$magePaths = @(
    'C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8.1 Tools\mage.exe',
    'C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\mage.exe'
)
$mage = $magePaths | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

foreach ($pair in $sources.GetEnumerator()) {
    $src = $pair.Value
    if (-not (Test-Path $src)) {
        Write-Host "SKIP $($pair.Key): source missing ($src). Build it first." -ForegroundColor Yellow
        continue
    }
    $dst = New-Item -ItemType Directory -Path (Join-Path $stage $pair.Key) -Force
    # Copy everything, drop debug symbols and stale backups to keep the package lean.
    $rc = & robocopy $src $dst.FullName /E /NP /NFL /NDL /XF '*.pdb' '*.bak' '*.bak*'
    if ($LASTEXITCODE -ge 8) { throw "robocopy failed for $($pair.Key) (exit $LASTEXITCODE)" }
    Write-Host "Staged $($pair.Key) ($([math]::Round((Get-ChildItem $dst.FullName -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB, 1)) MB)" -ForegroundColor Green

    # Re-sign manifests if mage is present and the source .vsto exists (keeps hashes in sync).
    if ($mage) {
        $vsto = Get-ChildItem -Path $dst.FullName -Filter *.vsto | Select-Object -First 1
        $dllManifest = Get-ChildItem -Path $dst.FullName -Filter *.dll.manifest | Select-Object -First 1
        if ($vsto -and $dllManifest -and (Test-Path -LiteralPath $CertPfx)) {
            & $mage -Sign $dllManifest.FullName -CertFile $CertPfx 2>&1 | Out-Null
            & $mage -Sign $vsto.FullName -CertFile $CertPfx 2>&1 | Out-Null
            Write-Host "  Re-signed manifests with $CertPfx" -ForegroundColor Gray
        }
    }
}

# ---- 3. certs + legal ----
$certDst = New-Item -ItemType Directory -Path (Join-Path $stage 'certs') -Force
if (Test-Path -LiteralPath $CertPfx) {
    Copy-Item -LiteralPath $CertPfx -Destination (Join-Path $certDst.FullName (Split-Path $CertPfx -Leaf))
    Write-Host "Bundled certificate: $(Split-Path $CertPfx -Leaf)" -ForegroundColor Green
} else {
    Write-Host "WARNING: certificate not found at $CertPfx - manifests will be untrusted on target." -ForegroundColor Yellow
}

$legalDst = New-Item -ItemType Directory -Path (Join-Path $stage 'Legal') -Force
foreach ($f in @('EULA.md', 'PRIVACY.md', 'THIRD_PARTY_NOTICES.md')) {
    $ls = Join-Path $repoRoot "Legal\$f"
    if (Test-Path $ls) { Copy-Item $ls $legalDst.FullName }
}
Write-Host 'Bundled Legal docs.' -ForegroundColor Green

# ---- 4. installer templates ----
$tmpl = Join-Path $scriptRoot 'install-templates'
foreach ($f in @('register.ps1', 'install.bat', 'uninstall.bat', 'README.txt')) {
    Copy-Item (Join-Path $tmpl $f) $stage
}
Write-Host 'Bundled installer templates.' -ForegroundColor Green

# ---- 5. portable ZIP ----
if (-not $NoZip) {
    $zip = Join-Path $out 'ExcelAIVSTOAddIn-Portable.zip'
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip -CompressionLevel Optimal
    Write-Host "ZIP  -> $zip ($([math]::Round((Get-Item $zip).Length / 1MB, 1)) MB)" -ForegroundColor Green
}

# ---- 6. Inno Setup ----
if (-not $NoInno) {
    $iss = Join-Path $out 'ExcelAIVSTOAddIn.iss'
    $issText = @"
; Inno Setup script for SmartPilot for Office (per-user, no admin).
; Generated by Tools\package-addin.ps1. Compile with ISCC.exe.
[Setup]
AppName=SmartPilot for Office
AppVersion=1.0.0.0
AppPublisher=SmartPilot
DefaultDirName={localappdata}\Programs\ExcelAIVSTOAddIn
DefaultGroupName=SmartPilot for Office
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline
OutputDir=$out
OutputBaseFilename=ExcelAIVSTOAddIn-setup
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayName=SmartPilot for Office
DisableProgramGroupPage=yes

[Files]
Source: "$stage\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Run]
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\register.ps1"""; Description: "Register Office add-ins"; Flags: runhidden

[UninstallRun]
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\register.ps1"" -Unregister"; Flags: runhidden

[Messages]
SetupAppTitle=SmartPilot for Office Setup
"@
    [System.IO.File]::WriteAllText($iss, $issText, [System.Text.Encoding]::UTF8)
    Write-Host "ISS  -> $iss" -ForegroundColor Green

    $iscc = Get-Command iscc.exe -ErrorAction SilentlyContinue
    if ($iscc) {
        & $iscc.FullName $iss 2>&1 | ForEach-Object { Write-Host "  iscc: $_" }
        $setup = Join-Path $out 'ExcelAIVSTOAddIn-setup.exe'
        if (Test-Path $setup) { Write-Host "SETUP -> $setup ($([math]::Round((Get-Item $setup).Length / 1MB, 1)) MB)" -ForegroundColor Green }
    } else {
        Write-Host "iscc.exe not found - produced $iss only. Install Inno Setup 6 (https://jrsoftware.org/isdl.php) and run: iscc `"$iss`"" -ForegroundColor Yellow
    }
}

Write-Host ''
Write-Host 'Package build finished.' -ForegroundColor Cyan
