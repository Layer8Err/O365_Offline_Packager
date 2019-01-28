## Decompress Office and install from local folder

## Get relative file paths
$rootFolder = $PSScriptRoot
if ($rootFolder.Length -lt 1){
    $rootFolder = $pwd.Path
    if ($rootFolder -match "Microsoft.PowerShell"){
        $rootFolder = $rootFolder.Split('::')[2]
    }
}
$zipfile = $rootFolder + '\CompressedOfficeInstaller.zip'
$destinationRoot = $env:TMP
$destinationfolder = $destinationRoot + '\MSOffice'
## Decompress into local temp folder
Add-Type -assembly 'system.io.compression.filesystem'
Write-Host "Extracting $zipfile to $dstinationfolder ..." -ForegroundColor Cyan
[io.compression.zipfile]::ExtractToDirectory($zipfile, $destinationfolder)
## Call installer from temp folder
Set-Location $destinationfolder
Write-Host "Installing from $destinationfolder ..." -ForegroundColor Cyan
cmd /c "Install365.bat"
