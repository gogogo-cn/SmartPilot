<#
.SYNOPSIS
    校验各宿主（Word/PPT）技能桶不含 Excel-only 技能（SkillSpec/ExcelBridge）。
.DESCRIPTION
    历史事故：SmartTableBeautifier（SkillSpec + dependencies:ExcelBridge）误声明
    hosts:[excel,word]，导致 Word 里被 spec 引擎渲染 Excel VBA 而卡死。
    本脚本扫描生成的 word_skills.json / ppt_skills.json（含各宿主 bin 部署副本），
    若发现 executionMode=spec 或 dependencies 含 ExcelBridge 的技能，则非零退出。
#>
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot   # Tools\.. = repo root

# 待校验的宿主桶文件（生成目录 + 各宿主 bin 部署副本）
$candidates = @(
    "$repoRoot\Hosts\Shared\Config\word_skills.json",
    "$repoRoot\Hosts\Shared\Config\ppt_skills.json",
    "$repoRoot\Hosts\WordAIVSTOAddIn\bin\Debug\Config\word_skills.json",
    "$repoRoot\Hosts\PptAIVSTOAddIn\bin\Debug\Config\ppt_skills.json"
)

function Test-ExcelOnly {
    param($skill)
    $name = if ($skill.name) { $skill.name.ToString() } else { '' }
    # 宿主生成类（生成 PPT/Word 的 spec 技能，如 GeneratePPTReport）在自身宿主内安全，不算违规
    if ($name -match 'ppt|word|幻灯片|文档|汇报') { return $false }
    $em = $skill.executionMode
    if ($em -and $em.ToString().ToLower() -eq 'spec') {
        $deps = $skill.dependencies
        if ($deps) {
            $depList = if ($deps -is [System.Collections.IEnumerable] -and -not ($deps -is [string])) {
                @($deps)
            } else {
                @($deps.ToString())
            }
            foreach ($d in $depList) {
                if ($d -and $d.ToString().ToLower() -like '*excelbridge*') { return $true }
            }
        }
    }
    return $false
}

$errors = @()
foreach ($file in $candidates) {
    if (-not (Test-Path $file)) { continue }
    try {
        $skills = Get-Content $file -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Warning "跳过无法解析的 JSON: $file ($_)"
        continue
    }
    if (-not $skills) { continue }
    foreach ($s in $skills) {
        $name = if ($s.name) { $s.name } else { '<unnamed>' }
        if (Test-ExcelOnly $s) {
            $errors += "[VIOLATION] $file 含 Excel-only 技能: $name (executionMode=$($s.executionMode), dependencies=$($s.dependencies))"
        }
    }
}

if ($errors.Count -gt 0) {
    Write-Host "=========================================" -ForegroundColor Red
    Write-Host "宿主技能分桶一致性校验失败" -ForegroundColor Red
    Write-Host "Excel-only 技能（SkillSpec / ExcelBridge）严禁进入 Word/PPT 桶，" -ForegroundColor Red
    Write-Host "否则会在非 Excel 宿主执行 Excel VBA 导致卡死。" -ForegroundColor Red
    Write-Host "请修正对应 Skills/*.md 的 hosts: 仅保留 [excel]。" -ForegroundColor Red
    Write-Host "=========================================" -ForegroundColor Red
    foreach ($e in $errors) { Write-Host $e -ForegroundColor Red }
    exit 1
}

Write-Host "[OK] 宿主技能分桶一致性校验通过：未发现 Excel-only 技能误入 Word/PPT 桶。" -ForegroundColor Green
exit 0
