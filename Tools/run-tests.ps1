# ============================================================
# Tools/run-tests.ps1 — 本地一键运行全部自动化测试
# 与 CI (build.yml) 的 test job 保持一致：
#   1) node:test 前端/服务端套件
#   2) VB CoreUtils 单元测试（dotnet run）
# 任一失败即退出码非 0。
# ============================================================
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

$failed = $false

Write-Host ""
Write-Host "=============================================="
Write-Host "  1/2  JS Unit Tests (node:test)"
Write-Host "=============================================="
Push-Location $root
node Tests/run-tests.js
if ($LASTEXITCODE -ne 0) {
    Write-Host "JS TEST FAILURES DETECTED" -ForegroundColor Red
    $failed = $true
} else {
    Write-Host "All JS test suites passed." -ForegroundColor Green
}
Pop-Location

Write-Host ""
Write-Host "=============================================="
Write-Host "  2/2  VB Core Unit Tests"
Write-Host "=============================================="
Push-Location $root
dotnet run --project Tests\VbTests\ExcelMcpCoreUtils.Tests.vbproj -c Release
if ($LASTEXITCODE -ne 0) {
    Write-Host "VB CORE TEST FAILURES DETECTED" -ForegroundColor Red
    $failed = $true
} else {
    Write-Host "All VB core tests passed." -ForegroundColor Green
}
Pop-Location

Write-Host ""
if ($failed) {
    Write-Host "RESULT: FAIL (some tests failed)" -ForegroundColor Red
    exit 1
}
Write-Host "RESULT: ALL TESTS PASSED" -ForegroundColor Green
exit 0