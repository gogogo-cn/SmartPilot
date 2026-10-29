# Excel VSTO Add-in Startup Diagnostic
param(
    [string]$AddInName = 'ExcelAIVSTOAddIn',
    [string]$BuildConfig = 'Debug'
)

$ErrorActionPreference = 'Continue'
$manifestPath = "file:///C:/Users/Administrator/Desktop/addin/ExcelAIVSTOAddIn/bin/$BuildConfig/ExcelAIVSTOAddIn.vsto|vstolocal"
$manifestFile = "C:\Users\Administrator\Desktop\addin\ExcelAIVSTOAddIn\bin\$BuildConfig\ExcelAIVSTOAddIn.vsto"

Write-Host "=== Excel VSTO Add-in Diagnostic: $AddInName ===" -ForegroundColor Cyan

# 1. Registry check
$regPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\$AddInName"
if (Test-Path $regPath) {
    $props = Get-ItemProperty -Path $regPath
    Write-Host "Registry entry exists:" -ForegroundColor Green
    Write-Host "  Manifest: $($props.Manifest)"
    Write-Host "  FriendlyName: $($props.FriendlyName)"
    Write-Host "  Description: $($props.Description)"
    Write-Host "  LoadBehavior: $($props.LoadBehavior)"
    if ($props.LoadBehavior -ne 3) {
        Write-Host "  [FIX] LoadBehavior is not 3, setting to 3..." -ForegroundColor Yellow
        Set-ItemProperty -Path $regPath -Name LoadBehavior -Value 3
    }
    if ($props.Manifest -ne $manifestPath) {
        Write-Host "  [FIX] Manifest path mismatch, correcting to $BuildConfig ..." -ForegroundColor Yellow
        Set-ItemProperty -Path $regPath -Name Manifest -Value $manifestPath
    }
} else {
    Write-Host "Registry entry missing, creating..." -ForegroundColor Yellow
    New-Item -Path $regPath -Force | Out-Null
    Set-ItemProperty -Path $regPath -Name Manifest -Value $manifestPath
    Set-ItemProperty -Path $regPath -Name FriendlyName -Value $AddInName
    Set-ItemProperty -Path $regPath -Name Description -Value "Excel AI VSTO Add-in"
    Set-ItemProperty -Path $regPath -Name LoadBehavior -Value 3
}

# 2. Resiliency disabled items
$resiliencyPaths = @(
    "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems",
    "HKCU:\Software\Microsoft\Office\15.0\Excel\Resiliency\DisabledItems"
)
foreach ($rp in $resiliencyPaths) {
    if (Test-Path $rp) {
        Write-Host "Found disabled items registry: $rp" -ForegroundColor Red
        $values = Get-ItemProperty -Path $rp
        $values.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
            Write-Host "  Disabled: $($_.Name) = $($_.Value)" -ForegroundColor Red
        }
        Write-Host "  [FIX] Removing disabled items..." -ForegroundColor Yellow
        Remove-Item -Path $rp -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        Write-Host "No disabled items: $rp" -ForegroundColor Green
    }
}

# 3. File existence
$files = @(
    $manifestFile,
    "C:\Users\Administrator\Desktop\addin\ExcelAIVSTOAddIn\bin\$BuildConfig\ExcelAIVSTOAddIn.dll",
    "C:\Users\Administrator\Desktop\addin\ExcelAIVSTOAddIn\bin\$BuildConfig\ExcelAIVSTOAddIn.dll.manifest",
    "C:\Users\Administrator\Desktop\addin\ExcelAIVSTOAddIn\bin\$BuildConfig\ExcelAIVSTOAddIn.WpfSidebar.dll"
)
foreach ($f in $files) {
    if (Test-Path $f) {
        $fi = Get-Item $f
        Write-Host "File exists: $f ($($fi.LastWriteTime))" -ForegroundColor Green
    } else {
        Write-Host "File missing: $f" -ForegroundColor Red
    }
}

# 4. Certificate check
$certPath = "C:\Users\Administrator\Desktop\addin\ExcelAIVSTOAddIn\ExcelAIVSTOAddIn_TemporaryKey.pfx"
if (Test-Path $certPath) {
    Write-Host "Temp certificate exists: $certPath" -ForegroundColor Green
} else {
    Write-Host "Temp certificate not found: $certPath" -ForegroundColor Yellow
}

Write-Host "`n=== Diagnostic complete ===" -ForegroundColor Cyan
Write-Host "Next: Close all Excel processes, then restart Excel." -ForegroundColor White
Write-Host "If still not starting, manually check: File -> Options -> Add-ins -> Manage COM Add-ins -> Go, and make sure $AddInName is checked." -ForegroundColor White
