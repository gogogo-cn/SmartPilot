# Install VSTO temporary key certificate to Trusted Publishers
$pfxPath = "c:\Users\Administrator\Desktop\addin\ExcelAIVSTOAddIn\ExcelAIVSTOAddIn_TemporaryKey.pfx"

$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import($pfxPath, "", "DefaultKeySet")

Write-Host "Certificate: $($cert.Subject)"
Write-Host "Thumbprint: $($cert.Thumbprint)"

# Check TrustedPublisher - CurrentUser
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("TrustedPublisher", "CurrentUser")
$store.Open("ReadWrite")

$found = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if ($found) {
    Write-Host "Already trusted in CurrentUser\TrustedPublisher"
} else {
    $store.Add($cert)
    Write-Host "Installed to CurrentUser\TrustedPublisher"
}
$store.Close()

# Also add to My store (Personal)
$myStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$myStore.Open("ReadWrite")
$found2 = $myStore.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if ($found2) {
    Write-Host "Already in CurrentUser\My"
} else {
    $myStore.Add($cert)
    Write-Host "Installed to CurrentUser\My"
}
$myStore.Close()

Write-Host "Done."
