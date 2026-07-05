# Signs a built DLL with the Excel Schedule Importer code-signing certificate.
# Called automatically by the SignAssembly MSBuild target after every build.
# Safe no-op if setup-signing.ps1 has not been run yet (build still succeeds,
# just produces an unsigned DLL as before).

param([Parameter(Mandatory = $true)][string]$Path)

$subject = 'CN=Excel Schedule Importer - Rocky Point Engineering'
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
        Where-Object { $_.Subject -eq $subject } | Select-Object -First 1

if (-not $cert) {
    Write-Host "SignAssembly: no signing certificate found - run setup-signing.ps1 once. Skipping (DLL stays unsigned)."
    exit 0
}

# Try with a timestamp server first (keeps the signature valid after the cert
# expires); fall back to signing without one if offline / server unreachable.
try {
    $result = Set-AuthenticodeSignature -FilePath $Path -Certificate $cert `
        -TimestampServer 'http://timestamp.digicert.com' -ErrorAction Stop
}
catch {
    $result = $null
}

if (-not $result -or $result.Status -ne 'Valid') {
    try {
        $result = Set-AuthenticodeSignature -FilePath $Path -Certificate $cert -ErrorAction Stop
    }
    catch {
        Write-Host "SignAssembly: failed to sign $Path : $_"
        exit 0   # never fail the build over signing
    }
}

if ($result.Status -eq 'Valid') {
    Write-Host "Signed: $Path"
} else {
    Write-Host "SignAssembly: signature status '$($result.Status)' - $($result.StatusMessage)"
}
