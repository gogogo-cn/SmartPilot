# Independent license test server. Not started by the Office add-in.
# Listens on http://127.0.0.1:9530 only. Domestic checkout stays in simulate mode.
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

function Test-Health {
    try {
        $r = Invoke-WebRequest -UseBasicParsing 'http://127.0.0.1:9530/health' -TimeoutSec 2
        return $r.StatusCode -eq 200
    } catch {
        return $false
    }
}

if (Test-Health) {
    Write-Host 'License test server is already running at http://127.0.0.1:9530'
    exit 0
}

$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    Write-Error 'Node.js 18+ is required. Install it, then run this script again. The add-in will not start this server.'
}

if (-not (Test-Path (Join-Path $Root '.env'))) {
    Copy-Item (Join-Path $Root '.env.example') (Join-Path $Root '.env')
}

if (-not (Test-Path (Join-Path $Root 'node_modules\express\package.json'))) {
    & npm install --omit=dev
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

$logDir = Join-Path $Root 'logs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$outLog = Join-Path $logDir 'test-server.out.log'
$errLog = Join-Path $logDir 'test-server.err.log'
$pidFile = Join-Path $logDir 'test-server.pid'

$env:BIND_HOST = '127.0.0.1'
$env:PORT = '9530'
$env:CN_PAY_PROVIDER = 'simulate'
if (-not $env:PUBLIC_BASE_URL) { $env:PUBLIC_BASE_URL = 'http://127.0.0.1:9530' }

$proc = Start-Process -FilePath $node.Source -ArgumentList 'server.js' -WorkingDirectory $Root -WindowStyle Hidden -PassThru -RedirectStandardOutput $outLog -RedirectStandardError $errLog
Set-Content -Path $pidFile -Value $proc.Id -Encoding ascii

$ready = $false
for ($i = 0; $i -lt 20; $i++) {
    if ($proc.HasExited) { break }
    if (Test-Health) { $ready = $true; break }
    Start-Sleep -Milliseconds 250
}

if (-not $ready) {
    Write-Host "License test server failed to listen. See $errLog"
    exit 1
}

Write-Host "License test server started (pid=$($proc.Id))"
Write-Host 'http://127.0.0.1:9530   health: /health   checkout: simulate'
Write-Host 'This process is independent of the add-in. Stop it with .\stop-test.ps1'
