<#
  package-web.ps1 - Package the desktop (browser + local bridge) edition.
  Produces dist\SmartPilotWeb-Portable.zip containing:
    UI\, Tools\py-excel-bridge.py, Config\, start.bat, README.txt
  The bridge serves the UI itself on port 9527 (browser opens http://127.0.0.1:9527/).
#>
param([string]$OutputDir = '')
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
if (-not $OutputDir) { $OutputDir = Join-Path $repo 'dist' }

$stage = Join-Path $OutputDir 'SmartPilotWeb'
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Path $stage -Force | Out-Null

# 1. full web UI
Copy-Item (Join-Path $repo 'UI') (Join-Path $stage 'UI') -Recurse
Write-Host 'Staged UI' -ForegroundColor Green

# 2. python bridge (expects ..\UI and ..\Config relative to Tools\)
New-Item -ItemType Directory -Path (Join-Path $stage 'Tools') -Force | Out-Null
Copy-Item (Join-Path $repo 'Tools\py-excel-bridge.py') (Join-Path $stage 'Tools\py-excel-bridge.py')
Write-Host 'Staged bridge' -ForegroundColor Green

# 3. config (tool defs / templates; no secrets inside)
Copy-Item (Join-Path $repo 'Config') (Join-Path $stage 'Config') -Recurse
Write-Host 'Staged Config' -ForegroundColor Green

# 3b. skills definitions (bridge reads Skills\*.md to serve GetAllSkills)
$skillDir = Join-Path $stage 'Skills'
New-Item -ItemType Directory -Path $skillDir -Force | Out-Null
Get-ChildItem (Join-Path $repo 'Skills') -Filter *.md -File | ForEach-Object {
    Copy-Item $_.FullName (Join-Path $skillDir $_.Name)
}
Write-Host ('Staged Skills (' + (Get-ChildItem $skillDir -File).Count + ' md)') -ForegroundColor Green

# 4. launcher + readme + uninstaller
# .bat 转为 GBK 编码：cmd 默认按系统代码页(936)解码，UTF-8 中文提示会乱码。
$gbk = [System.Text.Encoding]::GetEncoding(936)
foreach ($pair in @(@('start-web.bat', 'start.bat'), @('uninstall-web.bat', 'uninstall-web.bat'), @('SmartPilot.vbs', 'SmartPilot.vbs'))) {
    $srcText = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot ('install-templates\' + $pair[0])), [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText((Join-Path $stage $pair[1]), $srcText, $gbk)
}
Copy-Item (Join-Path $PSScriptRoot 'install-templates\web-README.txt') (Join-Path $stage 'README.txt')
Copy-Item (Join-Path $PSScriptRoot 'install-templates\smartpilot.ico') (Join-Path $stage 'smartpilot.ico')

# 5. zip
$zip = Join-Path $OutputDir 'SmartPilotWeb-Portable.zip'
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip -CompressionLevel Optimal
$size = [math]::Round((Get-Item $zip).Length / 1MB, 1)
Write-Host ("ZIP -> $zip ($size MB)") -ForegroundColor Cyan
