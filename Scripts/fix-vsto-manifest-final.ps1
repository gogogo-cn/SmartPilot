param(
    [string]$ReleaseDir = "c:\Users\Administrator\Desktop\addin\ExcelAIVSTOAddIn\bin\Release",
    [string]$RefDir = "c:\Users\Administrator\Desktop\ExcelAIVSTOAddIn\bin\Release",
    [string]$PfxFile = "c:\Users\Administrator\Desktop\addin\ExcelAIVSTOAddIn\ExcelAIVSTOAddIn_TemporaryKey.pfx"
)

$ErrorActionPreference = "Stop"

$mageExe = "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8.1 Tools\mage.exe"
$appManifest = Join-Path $ReleaseDir "ExcelAIVSTOAddIn.dll.manifest"
$depManifest = Join-Path $ReleaseDir "ExcelAIVSTOAddIn.vsto"

function Get-FileHashBase64 {
    param([string]$FilePath, [string]$Algorithm = "SHA256")
    $hasher = [System.Security.Cryptography.HashAlgorithm]::Create($Algorithm)
    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $hashBytes = $hasher.ComputeHash($fileBytes)
    return [System.Convert]::ToBase64String($hashBytes)
}

function Remove-Signature {
    param([System.Xml.XmlDocument]$Xml, [System.Xml.XmlNamespaceManager]$Ns)
    $sigNodes = $Xml.SelectNodes("//dsig:Signature", $Ns)
    foreach ($sig in $sigNodes) {
        [void]$sig.ParentNode.RemoveChild($sig)
    }
    $pubIdNodes = $Xml.SelectNodes("//asmv2:publisherIdentity", $Ns)
    foreach ($n in $pubIdNodes) {
        [void]$n.ParentNode.RemoveChild($n)
    }
}

Write-Host "=== Step 1: Copy reference manifests ===" -ForegroundColor Cyan
Copy-Item (Join-Path $RefDir "ExcelAIVSTOAddIn.dll.manifest") $appManifest -Force
Copy-Item (Join-Path $RefDir "ExcelAIVSTOAddIn.vsto") $depManifest -Force
Write-Host "App manifest: $((Get-Item $appManifest).Length) bytes"
Write-Host ".vsto: $((Get-Item $depManifest).Length) bytes"

Write-Host ""
Write-Host "=== Step 2: Update application manifest file hashes ===" -ForegroundColor Cyan

$appXml = New-Object System.Xml.XmlDocument
$appXml.PreserveWhitespace = $true
$appXml.Load($appManifest)

$appNs = New-Object System.Xml.XmlNamespaceManager($appXml.NameTable)
$appNs.AddNamespace("asmv2", "urn:schemas-microsoft-com:asm.v2")
$appNs.AddNamespace("dsig", "http://www.w3.org/2000/09/xmldsig#")

Remove-Signature -Xml $appXml -Ns $appNs

$updatedCount = 0
$missingCount = 0

$depAsms = $appXml.SelectNodes("//asmv2:dependentAssembly[@dependencyType='install']", $appNs)
Write-Host "Found $($depAsms.Count) dependentAssembly elements"

foreach ($depAsm in $depAsms) {
    $codebase = $depAsm.GetAttribute("codebase")
    if (-not $codebase) { continue }
    $filePath = Join-Path $ReleaseDir $codebase
    
    if (Test-Path $filePath) {
        $fileSize = (Get-Item $filePath).Length
        $depAsm.SetAttribute("size", [string]$fileSize)
        
        $hashNode = $depAsm.SelectSingleNode("dsig:hash", $appNs)
        if ($hashNode) {
            $digestValueNode = $hashNode.SelectSingleNode("dsig:DigestValue", $appNs)
            if ($digestValueNode) {
                $digestMethodNode = $hashNode.SelectSingleNode("dsig:DigestMethod", $appNs)
                $alg = "SHA256"
                if ($digestMethodNode -and $digestMethodNode.GetAttribute("Algorithm") -match "sha1") {
                    $alg = "SHA1"
                }
                $newHash = Get-FileHashBase64 -FilePath $filePath -Algorithm $alg
                $digestValueNode.InnerText = $newHash
            }
        }
        $updatedCount++
    } else {
        Write-Host "  WARNING: File not found: $codebase"
        $missingCount++
    }
}

$fileNodes = $appXml.SelectNodes("//asmv2:file", $appNs)
Write-Host "Found $($fileNodes.Count) file elements"

foreach ($fileNode in $fileNodes) {
    $name = $fileNode.GetAttribute("name")
    if (-not $name) { continue }
    $filePath = Join-Path $ReleaseDir $name
    
    if (Test-Path $filePath) {
        $fileSize = (Get-Item $filePath).Length
        $fileNode.SetAttribute("size", [string]$fileSize)
        
        $hashNode = $fileNode.SelectSingleNode("dsig:hash", $appNs)
        if ($hashNode) {
            $digestValueNode = $hashNode.SelectSingleNode("dsig:DigestValue", $appNs)
            if ($digestValueNode) {
                $digestMethodNode = $hashNode.SelectSingleNode("dsig:DigestMethod", $appNs)
                $alg = "SHA256"
                if ($digestMethodNode -and $digestMethodNode.GetAttribute("Algorithm") -match "sha1") {
                    $alg = "SHA1"
                }
                $newHash = Get-FileHashBase64 -FilePath $filePath -Algorithm $alg
                $digestValueNode.InnerText = $newHash
            }
        }
        $updatedCount++
    } else {
        Write-Host "  WARNING: File not found: $name"
        $missingCount++
    }
}

Write-Host "Updated: $updatedCount, Missing: $missingCount"

$appXml.Save($appManifest)
Write-Host "App manifest saved (unsigned). Size: $((Get-Item $appManifest).Length)" -ForegroundColor Green

Write-Host ""
Write-Host "=== Step 3: Sign application manifest ===" -ForegroundColor Cyan
& $mageExe -Sign $appManifest -CertFile $PfxFile 2>&1
Write-Host "Sign app exit code: $LASTEXITCODE"
Write-Host "App manifest size after sign: $((Get-Item $appManifest).Length)"

Write-Host ""
Write-Host "=== Step 4: Update deployment manifest ===" -ForegroundColor Cyan

$depXml = New-Object System.Xml.XmlDocument
$depXml.PreserveWhitespace = $true
$depXml.Load($depManifest)

$depNs = New-Object System.Xml.XmlNamespaceManager($depXml.NameTable)
$depNs.AddNamespace("asmv1", "urn:schemas-microsoft-com:asm.v1")
$depNs.AddNamespace("asmv2", "urn:schemas-microsoft-com:asm.v2")
$depNs.AddNamespace("dsig", "http://www.w3.org/2000/09/xmldsig#")

Remove-Signature -Xml $depXml -Ns $depNs

$appManifestSize = (Get-Item $appManifest).Length
Write-Host "App manifest size (signed): $appManifestSize"

$depAsm = $depXml.SelectSingleNode("//asmv1:dependency/asmv1:dependentAssembly", $depNs)
if ($depAsm) {
    $depAsm.SetAttribute("size", [string]$appManifestSize)
    Write-Host "Updated .vsto size attribute"
    
    $hashNode = $depAsm.SelectSingleNode("dsig:hash", $depNs)
    if ($hashNode) {
        $digestValueNode = $hashNode.SelectSingleNode("dsig:DigestValue", $depNs)
        $digestMethodNode = $hashNode.SelectSingleNode("dsig:DigestMethod", $depNs)
        if ($digestValueNode -and $digestMethodNode) {
            $alg = "SHA256"
            if ($digestMethodNode.GetAttribute("Algorithm") -match "sha1") {
                $alg = "SHA1"
            }
            $newHash = Get-FileHashBase64 -FilePath $appManifest -Algorithm $alg
            $digestValueNode.InnerText = $newHash
            Write-Host "Updated .vsto hash ($alg)"
        }
    }
}

$depXml.Save($depManifest)
Write-Host ".vsto saved (unsigned). Size: $((Get-Item $depManifest).Length)"

Write-Host ""
Write-Host "=== Step 5: Sign deployment manifest ===" -ForegroundColor Cyan
& $mageExe -Sign $depManifest -CertFile $PfxFile 2>&1
Write-Host "Sign .vsto exit code: $LASTEXITCODE"

Write-Host ""
Write-Host "=== DONE ===" -ForegroundColor Green
Write-Host "App manifest: $((Get-Item $appManifest).Length) bytes"
Write-Host ".vsto: $((Get-Item $depManifest).Length) bytes"

Write-Host ""
Write-Host "Verifying VSTO elements in app manifest:" -ForegroundColor Cyan
$checkXml = New-Object System.Xml.XmlDocument
$checkXml.Load($appManifest)
$vstoNodes = $checkXml.GetElementsByTagName("*") | Where-Object { $_.NamespaceURI -match "vsto|vsta" }
Write-Host "  Found $($vstoNodes.Count) VSTO/VSTA elements"
