# ============================================================
# sync-ui-runtime.ps1 — Windows: UI/ -> all runtime dirs + bump ?v=
# ============================================================
# Equivalent to Scripts/sync-ui-runtime.sh (use when bash/WSL unavailable).
# From repo root:
#   powershell -File Scripts/sync-ui-runtime.ps1
# ============================================================
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Src = Join-Path $Root 'UI'
$Index = Join-Path $Src 'index.html'

$Stamp = Get-Date -Format 'yyyyMMddHHmm'
$Ver = "v$Stamp"
Write-Host "==> cache version: $Ver"

# Bump ?v= and __APP_VERSION__ in all source HTML (WebView2 caches by URL query)
Get-ChildItem -Path $Src -Filter '*.html' -File | ForEach-Object {
    $raw = [IO.File]::ReadAllText($_.FullName)
    $raw2 = [regex]::Replace($raw, '\?v=v[0-9]{12}', "?v=$Ver")
    $raw2 = [regex]::Replace($raw2, "window\.__APP_VERSION__\s*=\s*'v[0-9]{12}'", "window.__APP_VERSION__ = '$Ver'")
    if ($raw2 -ne $raw) {
        [IO.File]::WriteAllText($_.FullName, $raw2)
        Write-Host "==> bumped: $($_.Name)"
    }
}

$Targets = @(
    (Join-Path $Root 'bin\Debug\UI'),
    (Join-Path $Root 'bin\Release\UI'),
    (Join-Path $Root 'Hosts\PptAIVSTOAddIn\bin\Debug\UI'),
    (Join-Path $Root 'Hosts\WordAIVSTOAddIn\bin\Debug\UI')
)

foreach ($T in $Targets) {
    if (-not (Test-Path $T)) {
        Write-Host "    skip (missing): $T"
        continue
    }
    Copy-Item -Path (Join-Path $Src '*') -Destination $T -Recurse -Force
    Write-Host "    synced: $T"
}

Write-Host '==> verify'
$Check = @(
    'index.html',
    'js\chat-v2\agent-loop.js',
    'js\chat-v2\services\tool-registry.js',
    'js\chat-v2\services\tool-registry.util.js'
)
$Failed = 0
foreach ($T in $Targets) {
    if (-not (Test-Path $T)) { continue }
    foreach ($F in $Check) {
        $A = Join-Path $Src $F
        $B = Join-Path $T $F
        if (-not (Test-Path $B)) {
            Write-Host "    missing: $B"
            $Failed = 1
            continue
        }
        $ha = (Get-FileHash $A -Algorithm SHA256).Hash
        $hb = (Get-FileHash $B -Algorithm SHA256).Hash
        if ($ha -ne $hb) {
            Write-Host "    mismatch: $B"
            $Failed = 1
        }
    }
}

if ($Failed -ne 0) {
    Write-Host '==> FAIL: runtime dirs out of sync with source'
    exit 1
}
Write-Host "==> done. version $Ver"
