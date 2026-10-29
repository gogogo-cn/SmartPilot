<#
.SYNOPSIS
    单目标 UI 同步（默认 bin/Debug/UI）
.DESCRIPTION
    日常开发请优先: bash Scripts/sync-ui-runtime.sh（4 处运行时 + 统一 ?v=）。
    本脚本仅同步到单个目标，封装 Tools/sync-ui-to-debug.js。
.PARAMETER Stamp
    强制指定版本戳，例如 v2026072355
.PARAMETER DryRun
    仅预览，不实际写入
.EXAMPLE
    .\Tools\sync-ui-to-debug.ps1
    .\Tools\sync-ui-to-debug.ps1 -DryRun
    .\Tools\sync-ui-to-debug.ps1 -Stamp v2026072355
#>
[CmdletBinding()]
param(
    [string]$Stamp,
    [switch]$DryRun,
    [switch]$Prune
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$nodeExe = Get-Command node -ErrorAction SilentlyContinue

if (-not $nodeExe) {
    Write-Host '❌ 未找到 Node.js，请先安装 Node.js (https://nodejs.org/)' -ForegroundColor Red
    exit 1
}

$argsList = @()
if ($Stamp) { $argsList += @('--stamp', $Stamp) }
if ($DryRun) { $argsList += '--dry-run' }
if ($Prune) { $argsList += '--prune' }

Write-Host '🔧 ExcelAIVSTOAddIn D1 构建同步' -ForegroundColor Cyan
Write-Host "项目路径: $projectRoot" -ForegroundColor Gray
Write-Host "Node 版本: $(node -v)" -ForegroundColor Gray
if ($Stamp) { Write-Host "指定版本戳: $Stamp" -ForegroundColor Yellow }
if ($DryRun) { Write-Host '模式: dry-run（仅预览）' -ForegroundColor Yellow }
if ($Prune) { Write-Host 'prune: 删除 bin/Debug 独有文件' -ForegroundColor Yellow }
Write-Host ''

$runner = Join-Path $PSScriptRoot 'sync-ui-to-debug.js'
& node $runner @argsList

exit $LASTEXITCODE
