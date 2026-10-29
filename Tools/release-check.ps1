<#
  release-check.ps1 — 上线/发布前自动检查清单
  ---------------------------------------------------------------------------
  覆盖上线前必须确认的高风险项（详见 Docs/DEPLOY.md）：
    1. 签名证书不能是临时证书（否则最终用户安装被 SmartScreen/Office 拦截）
    2. 应用版本号（ApplicationVersion/ApplicationRevision）已递增
    3. 各宿主 bin\Release\UI 的 HTML 版本戳(?v=)已刷新且全局唯一
       —— 注意：Release 构建现已自动刷戳（SyncUIDebugStamp 放宽到 Release），
          但若手工拷贝/未走 MSBuild 仍需本脚本复核
    4. MCP 工具一致性（VB 注册名 = JSON 名称 = 前端显示名）
    5. 前端 lint 零错误
    6. 自动化测试全绿（可用 -SkipTests 跳过）
    7. 法务文档齐全（EULA / 隐私说明 / 第三方声明）
    8. 发布目录不含真实 API 密钥或 .pfx
    9. Excel / Word / PPT 版本号一致
   10. ClickOnce InstallUrl（正式发布必填；-Ci 下降为警告）

  用法：
    cd 仓库根
    Tools\release-check.ps1                # 正式发布：临时证书与空 InstallUrl 为阻断
    Tools\release-check.ps1 -SkipTests     # 跳过耗时的测试环节
    Tools\release-check.ps1 -Ci -SkipTests # 日常 CI：证书与 InstallUrl 记为警告

  退出码：发现 ERROR 项返回 1，否则 0（可接入 CI 作为发布闸门）。
#>

param(
    [string]$RepoRoot = (Resolve-Path "$PSScriptRoot\.."),
    [switch]$SkipTests,
    [switch]$Ci
)

$ErrorActionPreference = 'Continue'
$errors   = 0
$warnings = 0

function Write-Hdr($t) { Write-Host "`n=== $t ===" -ForegroundColor Cyan }
function Add-Err($m)   { Write-Host "  [ERROR] $m" -ForegroundColor Red;   $script:errors++ }
function Add-Warn($m)  { Write-Host "  [WARN]  $m" -ForegroundColor Yellow; $script:warnings++ }
function Add-Ok($m)    { Write-Host "  [OK]    $m" -ForegroundColor Green }
function Add-Info($m)  { Write-Host "  [INFO]  $m" -ForegroundColor White }
function Invoke-ExternalPs([string]$file) {
    # 必须用 Start-Process：若用 & powershell.exe，子进程 stdout 会进入本函数的
    # 成功流，调用方拿到的就不是纯退出码；再拿数组做 -ne 0 会恒为真。
    $candidates = @(
        (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'),
        (Join-Path $PSHOME 'powershell.exe'),
        (Join-Path $PSHOME 'pwsh.exe')
    )
    foreach ($exe in $candidates) {
        if ($exe -and (Test-Path -LiteralPath $exe)) {
            $p = Start-Process -FilePath $exe -ArgumentList @(
                '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $file
            ) -Wait -PassThru -NoNewWindow
            if ($null -eq $p -or $null -eq $p.ExitCode) { return 0 }
            return [int]$p.ExitCode
        }
    }
    Add-Warn "找不到 powershell.exe，无法执行 $file"
    return 1
}

# ---------------------------------------------------------------------------
Write-Host "ExcelAIVSTOAddIn Release Check" -ForegroundColor White
Write-Host "RepoRoot: $RepoRoot"

$vbproj = Join-Path $RepoRoot 'ExcelAIVSTOAddIn.vbproj'
if (-not (Test-Path $vbproj)) { Add-Err "找不到 $vbproj"; exit 1 }
$vbprojText = Get-Content $vbproj -Raw

# ---------------------------------------------------------------------------
Write-Hdr "1. 签名证书（禁止临时证书分发）"
$tempThumb = 'ED1AC12D8BB72993C13AFA541F279849D1A1DF27'
if ($vbprojText -match 'ManifestKeyFile>\s*([^<]+)') {
    $keyFile = $Matches[1].Trim()
    Add-Info "ManifestKeyFile = $keyFile"
    if ($keyFile -match 'TemporaryKey') {
        $msg = "仍在使用临时证书 ($keyFile)，最终用户安装会被 SmartScreen 拦截；请替换为正式 EV/OV 证书并更新 ManifestKeyFile / ManifestCertificateThumbprint（见 Docs/DEPLOY.md）"
        if ($Ci) { Add-Warn $msg } else { Add-Err $msg }
    }
    else { Add-Ok "ManifestKeyFile 非临时证书" }
}
if ($vbprojText -match 'ManifestCertificateThumbprint>\s*([^<]+)') {
    $thumb = $Matches[1].Trim()
    Add-Info "ManifestCertificateThumbprint = $thumb"
    if ($thumb -eq $tempThumb) {
        $msg = "指纹仍是开发机临时证书指纹，正式发布必须替换"
        if ($Ci) { Add-Warn $msg } else { Add-Err $msg }
    }
    else { Add-Ok "指纹非临时证书" }
}

# ---------------------------------------------------------------------------
Write-Hdr "2. 应用版本号（ClickOnce 更新依赖）"
if ($vbprojText -match 'ApplicationVersion>\s*([^<]+)') { Add-Info "ApplicationVersion = $($Matches[1].Trim())" }
if ($vbprojText -match 'ApplicationRevision>\s*([^<]+)') { Add-Info "ApplicationRevision = $($Matches[1].Trim())" }
Add-Warn "请人工确认：每次对外发布是否已递增 ApplicationVersion / ApplicationRevision，否则已装用户收不到更新"

# ---------------------------------------------------------------------------
Write-Hdr "3. 各宿主 bin\Release\UI HTML 版本戳 (?v=)"
$hostUiDirs = @(
    (Join-Path $RepoRoot 'bin\Release\UI'),
    (Join-Path $RepoRoot 'Hosts\WordAIVSTOAddIn\bin\Release\UI'),
    (Join-Path $RepoRoot 'Hosts\PptAIVSTOAddIn\bin\Release\UI')
)
foreach ($uiDir in $hostUiDirs) {
    if (-not (Test-Path $uiDir)) {
        Add-Warn "未找到 $uiDir —— 可能该宿主尚未 Release 构建，跳过戳校验"
        continue
    }
    $index = Join-Path $uiDir 'index.html'
    if (-not (Test-Path $index)) { Add-Warn "$uiDir 无 index.html，跳过"; continue }
    $c = Get-Content $index -Raw
    $m = [regex]::Matches($c, '\?v=([^"''\s>]+)')
    if ($m.Count -eq 0) { Add-Err "$index 无 ?v= 版本戳"; continue }
    $uniq = @($m | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    if ($uniq.Count -gt 1) { Add-Err "$index 的 ?v= 不一致: $($uniq -join ', ')（WebView2 会混用旧脚本）" }
    else { Add-Ok "$index 版本戳一致 = $($uniq[0])" }
}

# ---------------------------------------------------------------------------
Write-Hdr "4. MCP 工具一致性校验"
$consistency = Join-Path $RepoRoot 'Tools\check-tools-consistency.ps1'
if (Test-Path $consistency) {
    $consistencyCode = Invoke-ExternalPs $consistency
    if ($consistencyCode -ne 0) { Add-Err "check-tools-consistency.ps1 未通过（VB 注册名 / JSON / 前端显示名 不一致）" }
    else { Add-Ok "工具一致性校验通过" }
} else { Add-Warn "未找到 $consistency，跳过" }

# ---------------------------------------------------------------------------
Write-Hdr "5. 前端 lint"
$lint = Join-Path $RepoRoot 'Tools\lint.js'
if (Test-Path $lint) {
    $node = Get-Command node -ErrorAction SilentlyContinue
    if ($node) {
        & node $lint
        if ($LASTEXITCODE -ne 0) { Add-Err "Tools\lint.js 报告错误（语法/规则子集未通过）" }
        else { Add-Ok "前端 lint 通过" }
    } else { Add-Warn "PATH 中无 node，跳过 lint" }
} else { Add-Warn "未找到 $lint，跳过" }

# ---------------------------------------------------------------------------
Write-Hdr "6. 自动化测试"
if ($SkipTests) {
    Add-Info "已通过 -SkipTests 跳过；上线前请务必手动运行 Tools\run-tests.ps1"
} else {
    $rt = Join-Path $RepoRoot 'Tools\run-tests.ps1'
    if (Test-Path $rt) {
        $testCode = Invoke-ExternalPs $rt
        if ($testCode -ne 0) { Add-Err "run-tests.ps1 未全绿" }
        else { Add-Ok "自动化测试全绿" }
    } else { Add-Warn "未找到 $rt，跳过" }
}

# ---------------------------------------------------------------------------
Write-Hdr "7. 法务文档"
foreach ($rel in @('Legal\EULA.md', 'Legal\PRIVACY.md', 'Legal\THIRD_PARTY_NOTICES.md')) {
    $p = Join-Path $RepoRoot $rel
    if (-not (Test-Path $p)) { Add-Err "缺少 $rel" }
    else {
        Add-Ok $rel
        $txt = Get-Content $p -Raw -ErrorAction SilentlyContinue
        if ($txt -match '【运营方】') { Add-Warn "$rel 仍含占位符【运营方】，对外发布前换成营业执照主体" }
    }
}

# ---------------------------------------------------------------------------
Write-Hdr "8. 密钥与打包卫生"
$giPath = Join-Path $RepoRoot '.gitignore'
if (Test-Path $giPath) {
    $gi = Get-Content $giPath -Raw
    if ($gi -match '(?m)^\*\.pfx\s*$') { Add-Ok ".gitignore 已忽略 *.pfx" }
    else { Add-Err ".gitignore 未忽略 *.pfx，代码签名私钥可能被提交" }
} else { Add-Err "缺少 .gitignore" }

function Test-ShipKey([string]$text) {
    if ($text -match '"apiKey"\s*:\s*"([^"]*)"') {
        $k = $Matches[1].Trim()
        if ($k -and $k -ne 'local' -and $k -notmatch '\*\*\*\*' -and $k.Length -ge 8) { return $true }
    }
    return $false
}

$shipRoots = @(
    (Join-Path $RepoRoot 'bin\Release'),
    (Join-Path $RepoRoot 'publish'),
    (Join-Path $RepoRoot 'Hosts\WordAIVSTOAddIn\bin\Release'),
    (Join-Path $RepoRoot 'Hosts\PptAIVSTOAddIn\bin\Release')
)
foreach ($rootDir in $shipRoots) {
    if (-not (Test-Path $rootDir)) { continue }
    $isPublish = $rootDir -match '\\publish$'
    Get-ChildItem $rootDir -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -eq '.pfx' -or $_.Name -eq 'llm_config.json' } |
        ForEach-Object {
        if ($_.Extension -eq '.pfx') {
            Add-Err "输出目录含证书私钥: $($_.FullName)"
        } elseif (Test-ShipKey (Get-Content $_.FullName -Raw -ErrorAction SilentlyContinue)) {
            $msg = "明文 API 密钥: $($_.FullName)。运行时写给本机独立服务可以，但 ClickOnce 安装包不得包含该文件"
            if ($isPublish) { Add-Err $msg } else { Add-Warn $msg }
        }
    }
}
$srcCfg = Join-Path $RepoRoot 'Config\llm_config.json'
if ((Test-Path $srcCfg) -and (Test-ShipKey (Get-Content $srcCfg -Raw))) {
    Add-Err "源码 Config\llm_config.json 含真实密钥，禁止入库或打进安装包"
}

# ---------------------------------------------------------------------------
Write-Hdr "9. 三宿主版本号"
function Get-AppVer([string]$projPath) {
    if (-not (Test-Path $projPath)) { return $null }
    $t = Get-Content $projPath -Raw
    if ($t -match 'ApplicationVersion>\s*([^<]+)') { return $Matches[1].Trim() }
    return $null
}
$excelVer = Get-AppVer (Join-Path $RepoRoot 'ExcelAIVSTOAddIn.vbproj')
$wordVer = Get-AppVer (Join-Path $RepoRoot 'Hosts\WordAIVSTOAddIn\WordAIVSTOAddIn.vbproj')
$pptAsm = Join-Path $RepoRoot 'Hosts\PptAIVSTOAddIn\My Project\AssemblyInfo.vb'
$pptVer = $null
if (Test-Path $pptAsm) {
    $pt = Get-Content $pptAsm -Raw
    if ($pt -match 'AssemblyVersion\("([^"]+)"\)') { $pptVer = $Matches[1].Trim() }
}
Add-Info "Excel ApplicationVersion = $excelVer"
Add-Info "Word  ApplicationVersion = $wordVer"
Add-Info "PPT   AssemblyVersion    = $pptVer"
if ($excelVer -and $wordVer -and $pptVer -and ($excelVer -eq $wordVer) -and ($excelVer -eq $pptVer)) {
    Add-Ok "三宿主版本号一致 = $excelVer"
} else {
    Add-Err "三宿主版本号不一致（Excel=$excelVer Word=$wordVer PPT=$pptVer），发布前必须对齐"
}

# ---------------------------------------------------------------------------
Write-Hdr "10. ClickOnce InstallUrl"
foreach ($pair in @(
    @{ Name = 'Excel'; Path = (Join-Path $RepoRoot 'ExcelAIVSTOAddIn.vbproj') },
    @{ Name = 'Word';  Path = (Join-Path $RepoRoot 'Hosts\WordAIVSTOAddIn\WordAIVSTOAddIn.vbproj') }
)) {
    if (-not (Test-Path $pair.Path)) { Add-Warn "$($pair.Name) 工程不存在"; continue }
    $t = Get-Content $pair.Path -Raw
    if ($t -match '<InstallUrl>\s*([^<]*)\s*</InstallUrl>') {
        $url = $Matches[1].Trim()
        if ([string]::IsNullOrWhiteSpace($url)) {
            $msg = "$($pair.Name) InstallUrl 为空。正式分发前填公网 HTTPS 更新目录，否则已安装用户无法 ClickOnce 更新"
            if ($Ci) { Add-Warn $msg } else { Add-Err $msg }
        } else { Add-Ok "$($pair.Name) InstallUrl = $url" }
    } else {
        $msg = "$($pair.Name) 未声明 InstallUrl"
        if ($Ci) { Add-Warn $msg } else { Add-Err $msg }
    }
}
Add-Info "PowerPoint 宿主当前按 COM 加载项注册，不走 Excel/Word 的 ClickOnce InstallUrl"

# ---------------------------------------------------------------------------
Write-Hdr "发布前人工待办（脚本无法自动判定的项）"
Add-Info "a. 已为三宿主分别做 Release 构建（Excel / Word / Ppt）"
Add-Info "b. 已重新生成 Word/PPT 技能分桶：python Hosts\Shared\Config\gen_host_skills.py"
Add-Info "c. 已用 VS 发布向导生成 ClickOnce 安装包到 PublishUrl (bin\Release\)"
Add-Info "d. 已确认目标机 Python + pywin32（若启用 Python 数据分析模板）"
Add-Info "e. 已完全退出 Office 进程后再验证新 DLL 生效"

# ---------------------------------------------------------------------------
Write-Host "`n================ 结果 ================" -ForegroundColor White
Write-Host "ERROR: $errors   WARN: $warnings" -ForegroundColor $(if ($errors -gt 0) { 'Red' } else { 'Green' })
if ($errors -gt 0) { Write-Host "存在阻断项，禁止发布。" -ForegroundColor Red; exit 1 }
Write-Host "无阻断项，可进入发布流程（仍需完成上方人工待办）。" -ForegroundColor Green
exit 0
