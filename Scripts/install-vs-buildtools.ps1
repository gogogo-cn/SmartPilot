# VS 2022 Build Tools + Office Dev Tools Installation Script
# Fixes P0 build environment issue

$ErrorActionPreference = "Stop"

$installerPath = "$env:USERPROFILE\Downloads\vs_BuildTools.exe"

# Step 1: Download installer
if (-not (Test-Path $installerPath)) {
    Write-Host "[1/3] Downloading VS Build Tools installer..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri "https://aka.ms/vs/17/release/vs_buildtools.exe" -OutFile $installerPath
    Write-Host "  Download complete: $installerPath" -ForegroundColor Green
} else {
    Write-Host "[1/3] Installer already exists: $installerPath" -ForegroundColor Yellow
}

# Step 2: Install Office development tools
Write-Host "[2/3] Installing VS Build Tools + Office Dev Tools..." -ForegroundColor Cyan
Write-Host "  This may take 10-30 minutes, please wait..." -ForegroundColor Yellow

$args = @(
    "--add", "Microsoft.VisualStudio.Workload.Office",
    "--add", "Microsoft.Net.Component.4.7.2.TargetingPack",
    "--passive",
    "--wait",
    "--norestart",
    "--includeRecommended"
)

$process = Start-Process -FilePath $installerPath -ArgumentList $args -Wait -PassThru -NoNewWindow

if ($process.ExitCode -ne 0) {
    Write-Host "  VS Build Tools installation failed, exit code: $($process.ExitCode)" -ForegroundColor Red
    Write-Host "  Please try running this script as Administrator" -ForegroundColor Yellow
    exit $process.ExitCode
}

Write-Host "  Installation complete!" -ForegroundColor Green

# Step 3: Verify installation
Write-Host "[3/3] Verifying installation..." -ForegroundColor Cyan
$vstoolsPath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Microsoft\VisualStudio\v18.0\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets"

if (Test-Path $vstoolsPath) {
    Write-Host "  VSTO build targets ready!" -ForegroundColor Green
    Write-Host "  $vstoolsPath" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  P0 fix complete! Project can now be compiled." -ForegroundColor Green
} else {
    Write-Host "  WARNING: VSTO build targets not found" -ForegroundColor Yellow
    Write-Host "  Expected: $vstoolsPath" -ForegroundColor Gray
    Write-Host "  Please verify Office Dev Tools were installed correctly" -ForegroundColor Yellow
}
