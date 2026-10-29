<#
.SYNOPSIS
    ExcelAIVSTOAddIn D2 端到端测试矩阵入口
.DESCRIPTION
    运行 L1/L2 可自动化测试，生成 JSON/Markdown 报告。
    L3/L4/L5 需手动在 Excel + 浏览器环境中执行，详见 Tools/e2e-test-matrix.md。
.EXAMPLE
    .\Tools\run-e2e-tests.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$nodeExe = Get-Command node -ErrorAction SilentlyContinue

if (-not $nodeExe) {
    Write-Host 'ERROR: Node.js not found, please install Node.js first (https://nodejs.org/)' -ForegroundColor Red
    exit 1
}

Write-Host 'ExcelAIVSTOAddIn D2 end-to-end test matrix' -ForegroundColor Cyan
Write-Host "Project path: $projectRoot" -ForegroundColor Gray
Write-Host "Node version: $(node -v)" -ForegroundColor Gray
Write-Host ''

$runner = Join-Path $PSScriptRoot 'e2e-runner.js'
& node $runner

$exitCode = $LASTEXITCODE
if ($exitCode -eq 0) {
    Write-Host ''
    Write-Host 'All L1/L2 automated tests passed.' -ForegroundColor Green
    Write-Host '  L3/L4/L5 manual test matrix: see Tools/e2e-test-matrix.md' -ForegroundColor Gray
} else {
    Write-Host ''
    Write-Host "Tests failed, exit code: $exitCode" -ForegroundColor Red
}

exit $exitCode
