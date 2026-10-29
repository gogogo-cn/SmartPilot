<#
  build-release.ps1 - Build all three VSTO hosts in Release configuration.
  Serial builds (parallel MSBuild triggers MSB3113).
  PreBuildEvent is overridden with "exit 0" to skip taskkill of Office processes;
  PostBuildEvent (UI sync + manifest signing) is preserved.
#>
param([string]$Configuration = 'Release')
$ErrorActionPreference = 'Continue'
$msbuild = 'D:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe'
if (-not (Test-Path $msbuild)) { throw "MSBuild not found: $msbuild" }
$repo = Split-Path $PSScriptRoot -Parent

$projects = @(
    'ExcelAIVSTOAddIn.vbproj',
    'Hosts\WordAIVSTOAddIn\WordAIVSTOAddIn.vbproj',
    'Hosts\PptAIVSTOAddIn\PptAIVSTOAddIn.vbproj'
)
$failed = @()
foreach ($p in $projects) {
    $full = Join-Path $repo $p
    Write-Host ('=== BUILD ' + $Configuration + ': ' + $p + ' ===') -ForegroundColor Cyan
    & $msbuild $full ('/p:Configuration=' + $Configuration) /p:Platform=AnyCPU /t:Build '/p:PreBuildEvent="exit 0"' /v:m /nologo 2>&1 | Out-Host
    $code = $LASTEXITCODE
    Write-Host ('EXIT=' + $code) -ForegroundColor Yellow
    if ($code -ne 0) { $failed += $p }
}
if ($failed.Count -gt 0) { Write-Host ('BUILD_FAILED: ' + ($failed -join ', ')) -ForegroundColor Red; exit 1 }
Write-Host 'ALL_RELEASE_BUILDS_OK' -ForegroundColor Green
