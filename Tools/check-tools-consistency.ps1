# ============================================================
# check-tools-consistency.ps1 — MCP 工具定义一致性校验
# ============================================================
# 作用：检测工具定义是否同步，避免配置文件与后端代码不一致
#   ① 后端 VB 注册：Services/Mcp/ExcelMcpTools.vb
#   ② 配置文件：Config/excel_tools.json（单一数据源，含 displayName/intentKeywords/operationType 等元数据）
#   ③ 前端：tool-registry.js 不得再硬编码显示名映射（显示名统一由 excel_tools.json 下发）
#
# 用法：
#   pwsh Tools/check-tools-consistency.ps1
#   或在 PowerShell 中直接运行
#
# 退出码：0=一致  1=存在差异（可用于构建前拦截）
# ============================================================

param(
    [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
)

# 设置控制台编码为 UTF-8，避免中文乱码
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$exitCode = 0
$hasDiff  = $false   # 汇总各校验项的差异标志（顶部初始化，避免被后续赋值覆盖）

# --- 路径定义 ---
$vbFile    = Join-Path $ProjectRoot 'Services\Mcp\ExcelMcpTools.vb'
$vbFolder  = Join-Path $ProjectRoot 'Services\Mcp'
$jsonFile  = Join-Path $ProjectRoot 'Config\excel_tools.json'
$jsFile    = Join-Path $ProjectRoot 'UI\js\chat-v2\services\tool-registry.js'
$tsFile    = Join-Path $ProjectRoot 'UI\js\chat-v2\services\tool-registry.ts'

Write-Host ""
Write-Host "=== MCP Tools Consistency Check ===" -ForegroundColor Cyan
Write-Host "Project Root: $ProjectRoot"
Write-Host ""

# --- 校验文件存在 ---
foreach ($f in @($vbFile, $jsonFile, $jsFile)) {
    if (-not (Test-Path $f)) {
        Write-Host "[ERROR] File not found: $f" -ForegroundColor Red
        exit 1
    }
}

# ============================================================
# 1. 提取后端 VB 注册的工具名
# 工具定义已拆分为多个 Partial 文件（ExcelMcpTools.*.vb），
# 因此遍历 Services\Mcp 下所有 .vb 汇总 tools.Add(New McpTool("..." 注册。
# ============================================================
$vbTools = @()
$vbPattern = 'tools\.Add\(New McpTool\("([^"]+)"'
foreach ($vbPath in (Get-ChildItem -Path $vbFolder -Filter '*.vb' -File -ErrorAction SilentlyContinue)) {
    $vbContent = Get-Content $vbPath.FullName -Raw -Encoding UTF8
    $vbTools += [regex]::Matches($vbContent, $vbPattern) | ForEach-Object { $_.Groups[1].Value }
}
$vbTools = $vbTools | Sort-Object -Unique
Write-Host "[1/3] Backend VB registered tools: $($vbTools.Count)" -ForegroundColor Yellow

# ============================================================
# 2. 提取 excel_tools.json 的工具名
# ============================================================
$jsonRaw = Get-Content $jsonFile -Raw -Encoding UTF8
# 去除可能的 BOM
if ($jsonRaw[0] -eq [char]0xFEFF) {
    $jsonRaw = $jsonRaw.Substring(1)
}
$jsonObj = $jsonRaw | ConvertFrom-Json
$jsonTools = $jsonObj.tools | ForEach-Object { $_.name } | Sort-Object -Unique
Write-Host "[2/3] excel_tools.json tools:       $($jsonTools.Count)" -ForegroundColor Yellow

# ============================================================
# 3. 校验前端 tool-registry.js 不再硬编码显示名映射
# ============================================================
$jsContent = Get-Content $jsFile -Raw -Encoding UTF8
$hasHardcodedMap = $jsContent.Contains('EXCEL_TOOL_DISPLAY_NAMES')
if ($hasHardcodedMap) {
    $mapStatus = 'FOUND (BAD)'
} else {
    $mapStatus = 'removed (OK)'
}
Write-Host "[3/3] Frontend hard-coded display map: $mapStatus" -ForegroundColor Yellow

# 校验 TS 镜像不再被引入（未编译迁移产物；若被重新引入且含硬编码映射则判为越界）
$hasTsHardcodedMap = $false
if (Test-Path $tsFile) {
    $tsContent = Get-Content $tsFile -Raw -Encoding UTF8
    $hasTsHardcodedMap = $tsContent.Contains('EXCEL_TOOL_DISPLAY_NAMES')
    if ($hasTsHardcodedMap) {
        Write-Host "[3b] tool-registry.ts 存在且含硬编码映射 (BAD)" -ForegroundColor Yellow
    } else {
        Write-Host "[3b] tool-registry.ts 存在但无硬编码映射（OK）" -ForegroundColor Yellow
    }
}
Write-Host ""

# ============================================================
# 4. 校验 excel_tools.json 元数据完整性（单一数据源）
# ============================================================
$missingMeta = @()
foreach ($t in $jsonObj.tools) {
    if (-not $t.displayName -or -not $t.operationType -or $null -eq $t.isWrite -or $null -eq $t.isDangerous) {
        $missingMeta += $t.name
    }
}
if ($missingMeta.Count -gt 0) {
    $hasDiff = $true
    Write-Host "[DIFF] excel_tools.json 元数据不完整（缺 displayName/operationType/isWrite/isDangerous）:" -ForegroundColor Red
    foreach ($t in $missingMeta) { Write-Host "  - $t" -ForegroundColor Red }
} else {
    Write-Host "[4/4] excel_tools.json 元数据完整: $($jsonObj.tools.Count) tools" -ForegroundColor Green
}
Write-Host ""

# ============================================================
# 差异比对（用简单的数组比较，兼容 PowerShell 5.x）
# ============================================================

# VB 有但 json 没有
$vbOnlyInJson = $vbTools | Where-Object { $jsonTools -notcontains $_ }
if ($vbOnlyInJson.Count -gt 0) {
    $hasDiff = $true
    Write-Host "[DIFF] In VB but MISSING in excel_tools.json:" -ForegroundColor Red
    foreach ($t in $vbOnlyInJson) { Write-Host "  - $t" -ForegroundColor Red }
}

# json 有但 VB 没有
$jsonOnlyInVb = $jsonTools | Where-Object { $vbTools -notcontains $_ }
if ($jsonOnlyInVb.Count -gt 0) {
    $hasDiff = $true
    Write-Host "[DIFF] In excel_tools.json but MISSING in VB:" -ForegroundColor Red
    foreach ($t in $jsonOnlyInVb) { Write-Host "  - $t" -ForegroundColor Red }
}

# 前端仍存在硬编码显示名映射
if ($hasHardcodedMap) {
    $hasDiff = $true
    Write-Host "[DIFF] tool-registry.js 仍包含硬编码 EXCEL_TOOL_DISPLAY_NAMES，应删除（显示名由 excel_tools.json 统一下发）" -ForegroundColor Red
}

# 未编译的 .ts 镜像被重新引入且含硬编码映射
if ($hasTsHardcodedMap) {
    $hasDiff = $true
    Write-Host "[DIFF] 检测到 tool-registry.ts 镜像仍含硬编码 EXCEL_TOOL_DISPLAY_NAMES。显示名应由 excel_tools.json/桥接统一下发，请删除该镜像或移除其硬编码映射。" -ForegroundColor Red
}

# ============================================================
# 结果输出
# ============================================================
if ($hasDiff) {
    Write-Host ""
    Write-Host "[RESULT] INCONSISTENT! Fix per project_memory.md sync checklist." -ForegroundColor Red
    Write-Host "Fix: excel_tools.json 是唯一数据源，VB 兜底与前端显示名必须与其一致" -ForegroundColor Red
    $exitCode = 1
} else {
    Write-Host "[RESULT] CONSISTENT (excel_tools.json 为单一数据源)." -ForegroundColor Green
    Write-Host "  Backend VB:          $($vbTools.Count) tools" -ForegroundColor Green
    Write-Host "  excel_tools.json:    $($jsonTools.Count) tools" -ForegroundColor Green
    Write-Host "  Frontend:            无硬编码显示名映射" -ForegroundColor Green
}

# 显示版本信息
if ($jsonObj.version) {
    Write-Host ""
    Write-Host "excel_tools.json version: $($jsonObj.version) (lastUpdated: $($jsonObj.lastUpdated))" -ForegroundColor DarkGray
}

# ============================================================
# 5. PPT 工具一致性（Config/ppt_tools.json ↔ PptMcpTools*.vb Partial）
# ============================================================
$pptJsonFile = Join-Path $ProjectRoot 'Config\ppt_tools.json'
$pptVbMain   = Join-Path $ProjectRoot 'Services\Mcp\PptMcpTools.vb'
$pptVbFiles  = @(Get-ChildItem -Path $vbFolder -Filter 'PptMcpTools*.vb' -File -ErrorAction SilentlyContinue)

Write-Host ""
Write-Host "=== PPT Tools Consistency Check ===" -ForegroundColor Cyan

if (-not (Test-Path $pptJsonFile) -or -not (Test-Path $pptVbMain)) {
    Write-Host "[ERROR] PPT 校验所需文件缺失: $pptJsonFile / $pptVbMain" -ForegroundColor Red
    exit 1
}

# 5a. 解析 ppt_tools.json
$pptRaw = Get-Content $pptJsonFile -Raw -Encoding UTF8
if ($pptRaw[0] -eq [char]0xFEFF) { $pptRaw = $pptRaw.Substring(1) }
$pptObj = $pptRaw | ConvertFrom-Json
$pptJsonTools = @($pptObj.tools | ForEach-Object { $_.name } | Sort-Object -Unique)
Write-Host "[5a] ppt_tools.json tools: $($pptJsonTools.Count)" -ForegroundColor Yellow

# 5b. 元数据完整性
$pptMissingMeta = @()
foreach ($t in @($pptObj.tools)) {
    if (-not $t.name -or -not $t.displayName -or -not $t.operationType -or $null -eq $t.isWrite -or $null -eq $t.isDangerous) {
        $pptMissingMeta += $t.name
    }
    if ($t.name -and $t.name -notmatch '^ppt_') {
        $pptMissingMeta += "$($t.name): 不以 ppt_ 开头"
    }
}
if ($pptMissingMeta.Count -gt 0) {
    $hasDiff = $true
    Write-Host "[DIFF] ppt_tools.json 元数据不完整/前缀不规范:" -ForegroundColor Red
    foreach ($t in $pptMissingMeta) { Write-Host "  - $t" -ForegroundColor Red }
} else {
    Write-Host "[5b] ppt_tools.json 元数据完整、前缀规范: $($pptObj.tools.Count) tools" -ForegroundColor Green
}

# 5c. 反射路由：每个 ppt_ 工具在任意 PptMcpTools*.vb Partial 中有对应 Execute 方法
#    ToolNameToMethodName: ppt_analyze_data → ExecuteAnalyzeData
#    分步构建工具在 PptMcpTools.Project.vb，不可只扫主文件。
$pptVbMethods = @()
foreach ($pptPath in $pptVbFiles) {
    $pptVbContent = Get-Content $pptPath.FullName -Raw -Encoding UTF8
    $pptVbMethods += [regex]::Matches($pptVbContent, '(?:Private|Public|Protected)\s+(?:Shared\s+)?Function\s+(Execute\w+)\s*\(') |
        ForEach-Object { $_.Groups[1].Value }
}
$pptVbMethods = @($pptVbMethods | Sort-Object -Unique)
Write-Host "[5c] PptMcpTools*.vb Execute 方法: $($pptVbMethods.Count)（文件 $($pptVbFiles.Count)）" -ForegroundColor Yellow

$pptMissingMethods = @()
foreach ($t in $pptJsonTools) {
    $short = $t
    if ($short.StartsWith('ppt_')) { $short = $short.Substring(4) }
    $method = 'Execute' + (($short -split '_' | Where-Object { $_ } | ForEach-Object { $_.Substring(0,1).ToUpper() + $_.Substring(1) }) -join '')
    if ($pptVbMethods -notcontains $method) {
        $pptMissingMethods += "$t → $method"
    }
}
if ($pptMissingMethods.Count -gt 0) {
    $hasDiff = $true
    Write-Host "[DIFF] ppt_tools.json 工具缺少对应 Execute 方法（反射路由会失败）:" -ForegroundColor Red
    foreach ($m in $pptMissingMethods) { Write-Host "  - $m" -ForegroundColor Red }
} else {
    Write-Host "[5c] ppt_tools.json 工具均有对应 Execute 方法: $($pptJsonTools.Count) tools" -ForegroundColor Green
}

if ($pptObj.version) {
    Write-Host "ppt_tools.json version: $($pptObj.version) (lastUpdated: $($pptObj.lastUpdated))" -ForegroundColor DarkGray
}

# PPT 段若发现差异，必须写回退出码（此前只改 $hasDiff 导致「有 DIFF 仍绿」）
if ($hasDiff) { $exitCode = 1 }

exit $exitCode
