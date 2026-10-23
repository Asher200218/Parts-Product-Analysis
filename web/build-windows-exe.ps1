# Build a standalone "Parts Analysis.exe" from launch.ps1.
# RUN THIS ON A WINDOWS PC (right-click -> Run with PowerShell, or run in a PowerShell window).
# The resulting .exe needs NO Python and NO PowerShell-module install to RUN -
# it just needs Windows. (Building it once uses the ps2exe module from PSGallery.)

$ErrorActionPreference = 'Stop'
$here = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "First-time only: installing the 'ps2exe' build tool from PSGallery..."
    try {
        Install-Module ps2exe -Scope CurrentUser -Force -ErrorAction Stop
    } catch {
        Write-Host ""
        Write-Host "Could not auto-install ps2exe. Install it once manually, then re-run:"
        Write-Host "    Install-Module ps2exe -Scope CurrentUser -Force"
        Read-Host "Press Enter to close"
        exit 1
    }
}
Import-Module ps2exe

$in  = Join-Path $here 'launch.ps1'
$out = Join-Path $here 'Parts Analysis.exe'

Invoke-ps2exe -inputFile $in -outputFile $out `
    -title 'Parts Analysis' `
    -product 'Parts Analysis' `
    -description 'Local parts/products/workers data-entry website (offline)' `
    -company 'local' -version '1.0.0'

Write-Host ""
Write-Host "Done -> $out"
Write-Host "Keep the .exe in this 'web' folder (next to index.html) and double-click it to run."
Read-Host "Press Enter to close"
