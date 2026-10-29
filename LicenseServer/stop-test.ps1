# Stop only the license test server started by start-test.ps1.
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$pidFile = Join-Path $Root 'logs\test-server.pid'
if (-not (Test-Path $pidFile)) {
    Write-Host 'No test server pid file. Nothing started by start-test.ps1.'
    exit 0
}
$pidText = (Get-Content $pidFile -Raw).Trim()
$procId = 0
if (-not [int]::TryParse($pidText, [ref]$procId)) {
    Remove-Item $pidFile -Force
    Write-Host 'Pid file was invalid and has been removed.'
    exit 0
}
$proc = Get-Process -Id $procId -ErrorAction SilentlyContinue
if ($proc) {
    Stop-Process -Id $procId -Force
    Write-Host "Stopped license test server (pid=$procId)"
} else {
    Write-Host "Process $procId is not running."
}
Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
