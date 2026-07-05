# Excel Schedule Importer - first-time installer
# ------------------------------------------------
# Run this ONCE per PC. It detects your installed Revit versions (2024/2025/2026),
# downloads the matching add-in from the latest GitHub release, and registers it.
# After this, the add-in keeps itself up to date automatically on each Revit launch.
#
# Usage (PowerShell):
#   irm https://github.com/AW-Designs/Excel-Import-to-Revit-plugin/releases/latest/download/install.ps1 -OutFile "$env:TEMP\esi-install.ps1"; & "$env:TEMP\esi-install.ps1"
#
# No admin rights required - everything installs under your user profile.

$ErrorActionPreference = 'Stop'
$owner = 'AW-Designs'
$repo  = 'Excel-Import-to-Revit-plugin'

Write-Host "Excel Schedule Importer - installer" -ForegroundColor Cyan

# Detect which Revit versions are present (by their Addins folder).
$addinsBase = Join-Path $env:APPDATA 'Autodesk\Revit\Addins'
$years = @(2024, 2025, 2026) | Where-Object {
    (Test-Path (Join-Path $addinsBase $_)) -or
    (Test-Path "C:\Program Files\Autodesk\Revit $_")
}
if (-not $years) {
    # Fall back: install for all three so it's ready whenever Revit is installed.
    $years = @(2024, 2025, 2026)
    Write-Host "No Revit install detected - installing for 2024/2025/2026 anyway." -ForegroundColor Yellow
}

try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}

$manifest = @"
<?xml version="1.0" encoding="utf-8"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>Excel Schedule Importer</Name>
    <Assembly>ExcelScheduleImporter\ExcelScheduleImporter.dll</Assembly>
    <FullClassName>ExcelScheduleImporter.App</FullClassName>
    <ClientId>7c1f3c1e-9a44-4d6b-b1a2-3e5f0d8c2a91</ClientId>
    <VendorId>AWDS</VendorId>
    <VendorDescription>Excel Schedule Importer</VendorDescription>
  </AddIn>
</RevitAddIns>
"@

foreach ($year in $years) {
    $yy = ($year.ToString()).Substring(2)   # 2026 -> 26
    $asset = "ExcelScheduleImporter-R$yy.zip"
    $url = "https://github.com/$owner/$repo/releases/latest/download/$asset"
    Write-Host "`nRevit $year : downloading $asset ..." -ForegroundColor Cyan

    $dest = Join-Path $addinsBase $year
    $pluginDir = Join-Path $dest 'ExcelScheduleImporter'
    New-Item -ItemType Directory -Force $dest | Out-Null
    New-Item -ItemType Directory -Force $pluginDir | Out-Null

    $tmp = Join-Path $env:TEMP "$asset"
    try {
        Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing
    } catch {
        Write-Host "  Skipped Revit $year (no release asset yet or download failed)." -ForegroundColor Yellow
        continue
    }

    # Clear old payload, extract new one, write the manifest.
    Get-ChildItem $pluginDir -Recurse -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Expand-Archive -Path $tmp -DestinationPath $pluginDir -Force
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    Set-Content -Path (Join-Path $dest 'ExcelScheduleImporter.addin') -Value $manifest -Encoding UTF8

    Write-Host "  Installed to $pluginDir" -ForegroundColor Green
}

Write-Host "`nDone. Start Revit - the 'Import Excel Schedule' button is on the Add-Ins tab." -ForegroundColor Green
Write-Host "The first time, click 'Always Load' if Revit asks about an unsigned add-in." -ForegroundColor Green
