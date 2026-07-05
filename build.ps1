# Builds Excel Schedule Importer for Revit 2024, 2025 and 2026
# and deploys each build to %AppData%\Autodesk\Revit\Addins\<year>\
#
# Requires the .NET 8 SDK. If dotnet is not on PATH, this script
# also looks in %USERPROFILE%\.dotnet (per-user install location).

$ErrorActionPreference = 'Stop'

$dotnet = 'dotnet'
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue) -or -not (& dotnet --list-sdks 2>$null)) {
    $local = Join-Path $env:USERPROFILE '.dotnet\dotnet.exe'
    if (Test-Path $local) { $dotnet = $local }
    else { throw ".NET SDK not found. Install from https://aka.ms/dotnet/download" }
}

$proj = Join-Path $PSScriptRoot 'src\ExcelScheduleImporter\ExcelScheduleImporter.csproj'

foreach ($cfg in 'Release R24', 'Release R25', 'Release R26') {
    Write-Host "`n=== Building $cfg ===" -ForegroundColor Cyan
    & $dotnet build $proj -c $cfg
    if ($LASTEXITCODE -ne 0) { throw "Build failed for $cfg" }
}

Write-Host "`nAll three versions built and deployed to %AppData%\Autodesk\Revit\Addins\{2024,2025,2026}" -ForegroundColor Green
