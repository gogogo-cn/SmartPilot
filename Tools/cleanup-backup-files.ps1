# 清理备份文件脚本
# 用途：删除 bin/ 和 Hosts/*/bin/ 下的所有 .bak-* 文件
# 运行方式：PowerShell -ExecutionPolicy Bypass -File cleanup-backup-files.ps1

$ErrorActionPreference = "Continue"
$projectRoot = Split-Path $PSScriptRoot -Parent

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  SmartPilot 备份文件清理工具" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 统计变量
$totalDeleted = 0
$totalSize = 0
$patterns = @(
    "*.bak-*",
    "*.bak"
)

# 定义要扫描的目录
$targetDirs = @(
    (Join-Path $projectRoot "bin"),
    (Join-Path $projectRoot "Hosts")
)

foreach ($dir in $targetDirs) {
    if (-not (Test-Path $dir)) {
        Write-Host "[跳过] 目录不存在: $dir" -ForegroundColor Yellow
        continue
    }

    Write-Host "[扫描] $dir" -ForegroundColor Green
    
    foreach ($pattern in $patterns) {
        $files = Get-ChildItem -Path $dir -Recurse -Filter $pattern -File -ErrorAction SilentlyContinue
        
        foreach ($file in $files) {
            $fileSize = $file.Length
            $totalSize += $fileSize
            $totalDeleted++
            
            # 显示删除信息（精简输出）
            $relativePath = $file.FullName.Replace($projectRoot, "").TrimStart('\')
            Write-Host "  [删除] $relativePath ($([math]::Round($fileSize/1KB, 1)) KB)" -ForegroundColor Gray
            
            Remove-Item $file.FullName -Force
        }
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  清理完成" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  删除文件数: $totalDeleted" -ForegroundColor White
Write-Host "  释放空间: $([math]::Round($totalSize/1MB, 2)) MB" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "提示：建议将此脚本添加到 .gitignore 或构建后事件中自动执行" -ForegroundColor Yellow