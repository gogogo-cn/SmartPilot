# VSTO Manifest Generator & Signer v3
param(
    [string]$OutputDir = "",
    [string]$PfxFile = "",
    [string]$AssemblyName = "ExcelAIVSTOAddIn",
    [string]$Version = "1.0.0.0",
    [string]$FriendlyName = "SmartPilot Excel AI AddIn",
    [string]$Description = "SmartPilot Excel AI AddIn",
    [string]$EntryPointClass = "ExcelAIVSTOAddIn.ThisAddIn",
    [string]$OfficeApp = "Microsoft Office Excel",
    [string]$KeyName = "ExcelAIVSTOAddIn",
    [int]$LoadBehavior = 3,
    [switch]$NoSign
)

$ErrorActionPreference = "Continue"

if (-not $OutputDir) { $OutputDir = (Get-Location).Path }
$OutputDir = $OutputDir.Trim('"').Trim()
if (-not $PfxFile) { 
    $PfxPath = Join-Path (Split-Path $PSScriptRoot -Parent) "ExcelAIVSTOAddIn_TemporaryKey.pfx"
} else {
    $PfxPath = $PfxFile.Trim('"').Trim()
}
$OutputDir = (Resolve-Path $OutputDir).Path
$AppManifest = Join-Path $OutputDir "$AssemblyName.dll.manifest"
$DepManifest = Join-Path $OutputDir "$AssemblyName.vsto"

Write-Host "OutputDir: $OutputDir"
Write-Host "PfxPath: $PfxPath"

# ============================================
# Generate file elements for ClickOnce manifest using mage.exe
# ============================================
function Get-FileElements {
    param([string]$OutputDir)
    
    $magePaths = @(
        "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8.1 Tools\mage.exe",
        "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\mage.exe"
    )
    $mageExe = $null
    foreach ($p in $magePaths) {
        if (Test-Path $p) { $mageExe = $p; break }
    }
    
    if (-not $mageExe) {
        Write-Host "WARNING: mage.exe not found, using PowerShell hash generation"
        return Get-FileElementsPowerShell -OutputDir $OutputDir
    }
    
    $tempManifest = Join-Path $OutputDir "temp_app.manifest"
    Write-Host "  Generating temp manifest with mage.exe..."
    & $mageExe -New Application -Name "TempApp" -Version 1.0.0.0 -Processor msil -FromDirectory $OutputDir -ToFile $tempManifest -TrustLevel FullTrust | Out-Null
    
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.Load($tempManifest)
    
    $nsManager = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
    $nsManager.AddNamespace("asmv2", "urn:schemas-microsoft-com:asm.v2")
    $nsManager.AddNamespace("dsig", "http://www.w3.org/2000/09/xmldsig#")
    
    $fileList = New-Object System.Collections.ArrayList
    foreach ($fileNode in $xmlDoc.SelectNodes("//asmv2:dependentAssembly[@dependencyType='install']", $nsManager)) {
        $codebase = $fileNode.GetAttribute("codebase")
        $size = $fileNode.GetAttribute("size")
        $hashNode = $fileNode.SelectSingleNode("hash", $nsManager)
        if (-not $hashNode) {
            $hashNode = $fileNode.SelectSingleNode("*[local-name()='hash']")
        }
        if ($hashNode) {
            $digestValue = $hashNode.SelectSingleNode("dsig:DigestValue", $nsManager).InnerText
            if (-not $digestValue) {
                $digestValNode = $hashNode.SelectSingleNode("*[local-name()='DigestValue']")
                if ($digestValNode) { $digestValue = $digestValNode.InnerText }
            }
            $digestMethod = $hashNode.SelectSingleNode("dsig:DigestMethod", $nsManager).GetAttribute("Algorithm")
            if (-not $digestMethod) {
                $digestMethodNode = $hashNode.SelectSingleNode("*[local-name()='DigestMethod']")
                if ($digestMethodNode) { $digestMethod = $digestMethodNode.GetAttribute("Algorithm") }
            }
            $fileList.Add("<file name=`"$codebase`" size=`"$size`"><hash><dsig:Transforms><dsig:Transform Algorithm=`"urn:schemas-microsoft-com:HashTransforms.Identity`"/></dsig:Transforms><dsig:DigestMethod Algorithm=`"$digestMethod`"/><dsig:DigestValue>$digestValue</dsig:DigestValue></hash></file>") | Out-Null
        }
        Write-Host "  File: $codebase (size=$size)"
    }
    
    Remove-Item $tempManifest -Force
    
    return $fileList -join "`n"
}

# ============================================
# Fallback: Calculate SHA256 hash for a file using PowerShell
# ============================================
function Get-FileElementsPowerShell {
    param([string]$OutputDir)
    
    function Get-FileHashBase64 {
        param([string]$FilePath)
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
        $hashBytes = $sha256.ComputeHash($fileBytes)
        return [System.Convert]::ToBase64String($hashBytes)
    }
    
    $files = Get-ChildItem -Path $OutputDir -File -Recurse | Where-Object {
        $_.Extension -in @('.dll', '.json', '.md', '.js', '.css', '.html', '.png', '.jpg', '.jpeg', '.gif', '.svg', '.ico', '.py', '.so', '.dylib') -and
        $_.Name -notin @('generate-vsto-manifests.ps1', 'deploy-exclude.txt') -and
        $_.FullName -notmatch '[\\/]node_modules[\\/]' -and
        $_.FullName -notmatch '[\\/]__pycache__[\\/]' -and
        $_.FullName -notmatch '[\\/]\.git[\\/]'
    }
    
    $fileElements = @()
    foreach ($file in $files) {
        $relativePath = $file.FullName.Substring($OutputDir.Length + 1)
        $relativePath = $relativePath.Replace('\', '/')
        $size = $file.Length
        $hash = Get-FileHashBase64 -FilePath $file.FullName
        $fileElements += "<file name=`"$relativePath`" size=`"$size`"><hash><dsig:Transforms><dsig:Transform Algorithm=`"urn:schemas-microsoft-com:HashTransforms.Identity`"/></dsig:Transforms><dsig:DigestMethod Algorithm=`"http://www.w3.org/2000/09/xmldsig#sha256`"/><dsig:DigestValue>$hash</dsig:DigestValue></hash></file>"
        Write-Host "  File: $relativePath (size=$size)"
    }
    
    return $fileElements -join "`n"
}

# ============================================
# Extract public key token from PFX certificate
# ============================================
function Get-PublicKeyTokenFromPfx {
    param([string]$PfxPath)
    try {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($PfxPath, "", 
            [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable -bor 
            [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet)
        $pk = $cert.PrivateKey
        if (-not $pk) {
            $pk = $cert.GetRSAPrivateKey()
        }
        if ($pk -is [System.Security.Cryptography.RSA]) {
            $rsaParams = $pk.ExportParameters($false)
            [byte[]]$exponent = if ($rsaParams.Exponent.Length -eq 3) {
                $rsaParams.Exponent + 0x00
            } else { $rsaParams.Exponent }
            [byte[]]$modulus = $rsaParams.Modulus
            $blobLen = 8 + 12 + $exponent.Length + $modulus.Length
            $mem = New-Object System.IO.MemoryStream
            $bw = New-Object System.IO.BinaryWriter($mem)
            $bw.Write([byte]0x06); $bw.Write([byte]0x02); $bw.Write([byte]0x00); $bw.Write([byte]0x00)
            $bw.Write([int][System.BitConverter]::ToInt32([System.Text.Encoding]::ASCII.GetBytes("RSA1"), 0))
            $bw.Write($modulus.Length * 8)
            $bw.Write([int]0)
            $bw.Write($exponent.Length)
            $bw.Write($exponent)
            $bw.Write($modulus.Length)
            $bw.Write($modulus)
            $bw.Flush()
            [byte[]]$pubKeyBlob = $mem.ToArray()
            $bw.Close(); $mem.Close()
            $sha1 = [System.Security.Cryptography.SHA1]::Create()
            $hash = $sha1.ComputeHash($pubKeyBlob)
            [byte[]]$token = $hash[($hash.Length-8)..($hash.Length-1)]
            [Array]::Reverse($token)
            return [System.BitConverter]::ToString($token) -replace '-',''
        }
    } catch {
        Write-Host "WARNING: Could not extract publicKeyToken from PFX: $_"
    }
    return $null
}
$publicKeyToken = Get-PublicKeyTokenFromPfx -PfxPath $PfxPath
Write-Host "PublicKeyToken: $publicKeyToken"

# ============================================
# Generate Application Manifest XML
# ============================================
$pubKeyTokenAttr = if ($publicKeyToken) { "publicKeyToken=`"$publicKeyToken`"" } else { "" }
Write-Host "Generating file elements..."
$fileElements = Get-FileElements -OutputDir $OutputDir

$appXmlHead = @"
<?xml version="1.0" encoding="utf-8"?>
<asmv1:assembly xsi:schemaLocation="urn:schemas-microsoft-com:asm.v1 assembly.adaptive.xsd" 
  manifestVersion="1.0" 
  xmlns:asmv1="urn:schemas-microsoft-com:asm.v1" 
  xmlns="urn:schemas-microsoft-com:asm.v2" 
  xmlns:asmv2="urn:schemas-microsoft-com:asm.v2" 
  xmlns:asmv3="urn:schemas-microsoft-com:asm.v3" 
  xmlns:dsig="http://www.w3.org/2000/09/xmldsig#" 
  xmlns:co.v1="urn:schemas-microsoft-com:clickonce.v1" 
  xmlns:co.v2="urn:schemas-microsoft-com:clickonce.v2"
  xmlns:vstav3="urn:schemas-microsoft-com:vsta.v3"
  xmlns:vstov4="urn:schemas-microsoft-com:vsto.v4"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <asmv1:assemblyIdentity name="$AssemblyName.dll" version="$Version" language="neutral" processorArchitecture="msil" $pubKeyTokenAttr />
  <vstov4:entryPoints>
    <vstov4:entryPoint class="$EntryPointClass">
      <assemblyIdentity name="$AssemblyName" version="$Version" language="neutral" processorArchitecture="msil" $pubKeyTokenAttr />
    </vstov4:entryPoint>
  </vstov4:entryPoints>
  <vstav3:addIn 
    application="$OfficeApp" 
    loadBehavior="$LoadBehavior" 
    keyName="$KeyName"
    xmlns:vstav3="urn:schemas-microsoft-com:vsta.v3">
    <vstav3:friendlyName>$FriendlyName</vstav3:friendlyName>
    <vstav3:description>$Description</vstav3:description>
    <vstav3:formRegions />
  </vstav3:addIn>
"@

$appXmlTail = "</asmv1:assembly>"

$appXml = $appXmlHead + $fileElements + $appXmlTail

# ============================================
# Generate Deployment Manifest XML
# ============================================
$depXml = @"
<?xml version="1.0" encoding="utf-8"?>
<asmv1:assembly xsi:schemaLocation="urn:schemas-microsoft-com:asm.v1 assembly.adaptive.xsd" 
  manifestVersion="1.0" 
  xmlns:asmv1="urn:schemas-microsoft-com:asm.v1" 
  xmlns="urn:schemas-microsoft-com:asm.v2" 
  xmlns:asmv2="urn:schemas-microsoft-com:asm.v2" 
  xmlns:asmv3="urn:schemas-microsoft-com:asm.v3" 
  xmlns:dsig="http://www.w3.org/2000/09/xmldsig#" 
  xmlns:co.v1="urn:schemas-microsoft-com:clickonce.v1" 
  xmlns:co.v2="urn:schemas-microsoft-com:clickonce.v2"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <asmv1:assemblyIdentity name="$AssemblyName.vsto" version="$Version" language="neutral" processorArchitecture="msil" $pubKeyTokenAttr />
  <description asmv2:description="$Description" asmv2:publisher="$KeyName" asmv2:product="$FriendlyName" xmlns="urn:schemas-microsoft-com:asm.v1" />
  <deployment install="false" mapFileExtensions="true" trustURLParameters="false" />
  <dependency>
    <dependentAssembly dependencyType="install" codebase="$AssemblyName.dll.manifest" size="0">
      <assemblyIdentity name="$AssemblyName.dll" version="$Version" language="neutral" processorArchitecture="msil" $pubKeyTokenAttr />
    </dependentAssembly>
  </dependency>
</asmv1:assembly>
"@

# ============================================
# Sign a ClickOnce manifest using mage.exe (Microsoft's official tool)
# Use for: deployment manifests (.vsto, .application)
# ============================================
function Sign-DeploymentManifest {
    param([string]$ManifestPath, [string]$PfxPath)
    
    Write-Host "Signing (mage): $ManifestPath"
    
    $magePaths = @(
        "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8.1 Tools\mage.exe",
        "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\mage.exe"
    )
    $mageExe = $null
    foreach ($p in $magePaths) {
        if (Test-Path $p) { $mageExe = $p; break }
    }
    
    if (-not $mageExe) {
        throw "mage.exe not found. Install .NET Framework SDK."
    }
    
    Write-Host "  Using: $mageExe"
    $result = & $mageExe -Sign $ManifestPath -CertFile $PfxPath 2>&1
    $exitCode = $LASTEXITCODE
    
    if ($exitCode -ne 0) {
        throw "mage.exe failed (exit=${exitCode}): $result"
    }
    Write-Host "  Signed successfully!"
}

# ============================================
# Sign an application manifest (.dll.manifest) using mage.exe
# ============================================
function Sign-ApplicationManifest {
    param([string]$ManifestPath, [string]$PfxPath)
    
    Write-Host "Signing (mage): $ManifestPath"
    
    $magePaths = @(
        "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8.1 Tools\mage.exe",
        "C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\mage.exe"
    )
    $mageExe = $null
    foreach ($p in $magePaths) {
        if (Test-Path $p) { $mageExe = $p; break }
    }
    
    if (-not $mageExe) {
        throw "mage.exe not found. Install .NET Framework SDK."
    }
    
    Write-Host "  Using: $mageExe"
    $result = & $mageExe -Sign $ManifestPath -CertFile $PfxPath 2>&1
    $exitCode = $LASTEXITCODE
    
    if ($exitCode -ne 0) {
        throw "mage.exe failed (exit=${exitCode}): $result"
    }
    Write-Host "  Signed successfully!"
}

# ============================================
# Write application manifest (overwrite existing)
# ============================================
Write-Host "Writing application manifest..."
[System.IO.File]::WriteAllText($AppManifest, $appXml, [System.Text.Encoding]::UTF8)

$appManifestSize = (Get-Item $AppManifest).Length
Write-Host "Application manifest size: $appManifestSize"

# ============================================
# Sign manifests (skip if -NoSign)
# ============================================
if ($NoSign) {
    $depXml = @"
<?xml version="1.0" encoding="utf-8"?>
<asmv1:assembly xsi:schemaLocation="urn:schemas-microsoft-com:asm.v1 assembly.adaptive.xsd" 
  manifestVersion="1.0" 
  xmlns:asmv1="urn:schemas-microsoft-com:asm.v1" 
  xmlns="urn:schemas-microsoft-com:asm.v2" 
  xmlns:asmv2="urn:schemas-microsoft-com:asm.v2" 
  xmlns:asmv3="urn:schemas-microsoft-com:asm.v3" 
  xmlns:dsig="http://www.w3.org/2000/09/xmldsig#" 
  xmlns:co.v1="urn:schemas-microsoft-com:clickonce.v1" 
  xmlns:co.v2="urn:schemas-microsoft-com:clickonce.v2"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <asmv1:assemblyIdentity name="$AssemblyName.vsto" version="$Version" language="neutral" processorArchitecture="msil" $pubKeyTokenAttr />
  <description asmv2:description="$Description" asmv2:publisher="$KeyName" asmv2:product="$FriendlyName" xmlns="urn:schemas-microsoft-com:asm.v1" />
  <deployment install="false" mapFileExtensions="true" trustURLParameters="false" />
  <dependency>
    <dependentAssembly dependencyType="install" codebase="$AssemblyName.dll.manifest" size="$appManifestSize">
      <assemblyIdentity name="$AssemblyName.dll" version="$Version" language="neutral" processorArchitecture="msil" $pubKeyTokenAttr />
    </dependentAssembly>
  </dependency>
</asmv1:assembly>
"@
    Write-Host "Writing deployment manifest..."
    [System.IO.File]::WriteAllText($DepManifest, $depXml, [System.Text.Encoding]::UTF8)
    Write-Host ""
    Write-Host "Skipping signing (-NoSign specified)."
    Write-Host "============================================"
    Write-Host "SUCCESS: Unsigned manifests created."
    Write-Host "  $AppManifest"
    Write-Host "  $DepManifest"
    Write-Host "============================================"
    exit 0
}

Write-Host ""
Write-Host "Signing manifests..."
$signOk = $true

try {
    Sign-ApplicationManifest -ManifestPath $AppManifest -PfxPath $PfxPath
} catch {
    Write-Host "ERROR signing application manifest: $_"
    $signOk = $false
}

$xmlDoc = New-Object System.Xml.XmlDocument
$xmlDoc.Load($AppManifest)
$nsManager = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
$nsManager.AddNamespace("asmv1", "urn:schemas-microsoft-com:asm.v1")
$identityNode = $xmlDoc.SelectSingleNode("//asmv1:assemblyIdentity", $nsManager)
$actualPublicKeyToken = $identityNode.GetAttribute("publicKeyToken")
Write-Host "Application manifest publicKeyToken: $actualPublicKeyToken"

$depXml = @"
<?xml version="1.0" encoding="utf-8"?>
<asmv1:assembly xsi:schemaLocation="urn:schemas-microsoft-com:asm.v1 assembly.adaptive.xsd" 
  manifestVersion="1.0" 
  xmlns:asmv1="urn:schemas-microsoft-com:asm.v1" 
  xmlns="urn:schemas-microsoft-com:asm.v2" 
  xmlns:asmv2="urn:schemas-microsoft-com:asm.v2" 
  xmlns:asmv3="urn:schemas-microsoft-com:asm.v3" 
  xmlns:dsig="http://www.w3.org/2000/09/xmldsig#" 
  xmlns:co.v1="urn:schemas-microsoft-com:clickonce.v1" 
  xmlns:co.v2="urn:schemas-microsoft-com:clickonce.v2"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <asmv1:assemblyIdentity name="$AssemblyName.vsto" version="$Version" language="neutral" processorArchitecture="msil" publicKeyToken="$actualPublicKeyToken" />
  <description asmv2:description="$Description" asmv2:publisher="$KeyName" asmv2:product="$FriendlyName" xmlns="urn:schemas-microsoft-com:asm.v1" />
  <deployment install="false" mapFileExtensions="true" trustURLParameters="false" />
  <dependency>
    <dependentAssembly dependencyType="install" codebase="$AssemblyName.dll.manifest" size="0">
      <assemblyIdentity name="$AssemblyName.dll" version="$Version" language="neutral" processorArchitecture="msil" publicKeyToken="$actualPublicKeyToken" />
    </dependentAssembly>
  </dependency>
</asmv1:assembly>
"@

Write-Host "Writing deployment manifest..."
[System.IO.File]::WriteAllText($DepManifest, $depXml, [System.Text.Encoding]::UTF8)

try {
    Sign-DeploymentManifest -ManifestPath $DepManifest -PfxPath $PfxPath
} catch {
    Write-Host "ERROR signing deployment manifest: $_"
    $signOk = $false
}

Write-Host ""
Write-Host "============================================"
if ($signOk) {
    Write-Host "SUCCESS: Manifests created and signed."
} else {
    Write-Host "WARNING: Manifests created but signing failed."
    Write-Host "You may need to manually sign them or disable signing in .vbproj."
}
Write-Host "  $AppManifest"
Write-Host "  $DepManifest"
Write-Host "============================================"
