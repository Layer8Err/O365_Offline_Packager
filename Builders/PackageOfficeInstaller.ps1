## Package Office for distribution
###########################################################
# Installer packaging Script for creating a compressed
# Office 365 Pro Plus installation package
# 
# This tool will take downloaded install packages and
# compress them into a folder with install
# scripts for deployment. The created zip file:
# "CompressedOfficeInstaller.zip" can be re-deployed in
# orer to update the install package.
#
# All install files are placed in the
#  Office_365_CompressedInstall folder
#
###########################################################
Set-Location $PSScriptRoot
$thisDir = $PWD.Path
$dirarr = $thisDir.Split('\')
$upDir = ($dirarr | Select-Object -First ($dirArr.Length - 1)) -join '\'
$source = $thisdir + '\OfficeInstaller'
$destination = $thisdir + '\CompressedOfficeInstaller.zip'
Write-Host "Checking for existing CompressedOfficeInstaller.zip file..." -ForegroundColor Yellow
If(Test-path $destination) {
    Write-Host "Found existing zip file. Removing..." -ForegroundColor Cyan
    Remove-item $destination
}
Write-Host "Compressing OfficeInstaller contents (this may take a bit)..." -ForegroundColor Green
Add-Type -assembly 'system.io.compression.filesystem'
[io.compression.zipfile]::CreateFromDirectory($source, $destination)
Write-Host "... Done" -ForegroundColor White

# Create Office_365_CompressedInstall package
$installpkg = $upDir + '\Office_365_CompressedInstall'
Write-Host "Checking for existing Office_365_CompressedInstall folder..."
If(Test-Path $installpkg){
    Write-Host "Founding existing Compressed Install folder. Removing..." -ForegroundColor Cyan
    Remove-Item -Recurse -Path $installpkg -Force -Confirm:$false
}
Write-Host "Building distributable Office 365 install package..."
mkdir -Path $installpkg
Move-Item -Path $destination -Destination $installpkg\.
Copy-Item -Path ($thisDir + '\Decompress_install_Office.ps1') -Destination $installpkg\.
Copy-Item -Path ($thisDir + '\Install.bat') -Destination $installpkg\.

# Cleanup Leftover files from Build process
Write-Host "Cleaning up build files..."
if (Test-Path $source) {
    Write-Host "Removing downloaded install files..."
    Remove-Item -Recurse -Path $source -Force -Confirm:$false
}
if (Test-Path ($thisDir + '\*.exe')) {
    Write-Host "Removing downloaded Office deployment tool..."
    Remove-Item *.exe -Force
}
if (Test-Path ($thisDir + '\*.xml')) {
    Write-Host "Removing XML configuration file(s)..."
    Remove-Item *.xml -Force
}
Write-Host "... Done" -ForegroundColor White

PAUSE