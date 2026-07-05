# One-time setup: creates a self-signed code-signing certificate for
# Excel Schedule Importer and trusts it on THIS machine (current user).
#
# After this runs, build.ps1 automatically signs the DLL with this certificate
# (see sign.ps1 / the SignAssembly target in the .csproj). Because Revit's
# "Always Load" decision is tied to the certificate (not the file hash), you
# only need to click it once - future rebuilds keep loading silently.
#
# Re-run safely any time; it reuses the existing certificate if present.

$ErrorActionPreference = 'Stop'
$subject = 'CN=Excel Schedule Importer - Rocky Point Engineering'

$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
        Where-Object { $_.Subject -eq $subject } | Select-Object -First 1

if (-not $cert) {
    Write-Host "Creating self-signed code-signing certificate..." -ForegroundColor Cyan
    $cert = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $subject `
        -CertStoreLocation Cert:\CurrentUser\My `
        -NotAfter (Get-Date).AddYears(20) `
        -KeyExportPolicy Exportable `
        -KeyUsage DigitalSignature `
        -KeyAlgorithm RSA `
        -KeyLength 2048
} else {
    Write-Host "Certificate already exists (thumbprint $($cert.Thumbprint)) - reusing it." -ForegroundColor Cyan
}

# Trust it locally: self-signed, so the cert must be trusted both as a root
# and as a trusted publisher for Windows/Revit to accept it without prompting
# about the certificate itself (Revit's own "unsigned add-in" prompt is
# separate and still appears once, showing the publisher name from this cert).
#
# Import-Certificate pops a confirmation dialog for Root adds, which fails
# non-interactively ("UI is not allowed"). Adding via X509Store directly is
# the same underlying operation without that UI requirement.
foreach ($storeName in 'Root', 'TrustedPublisher') {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
        $storeName, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $store.Add($cert)
    $store.Close()
}

Write-Host "`nDone. Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green
Write-Host "Run build.ps1 to build a signed DLL, then in Revit choose 'Always Load'" -ForegroundColor Green
Write-Host "once - it will stick across future rebuilds since the certificate stays the same." -ForegroundColor Green
