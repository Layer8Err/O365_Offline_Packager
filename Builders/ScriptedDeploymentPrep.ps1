###########################################################
# Auto-Build Script for creating an Office 365 Pro Plus
# installation package
# 
# This tool will automatically download the
# officedeploymenttool from Microsoft and create an
# installation package ready for deployment this will
# ensure that the deployment package is the latest version.
#
###########################################################
## Assume that script is run from root dir
Set-Location $PSScriptRoot
$buildroot = $PWD.Path

## Download the officedeploymenttool from Microsoft
Write-Host "Beginning Office 365 setuptool download..."
$websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$url = "https://www.microsoft.com/en-us/download/details.aspx?id=49117" # Microsoft's landing page
$dlurl = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117" # Download url for setuptool (URL for "Download" button)
$cookies = $websession.Cookies.GetCookies($url)
$websession.Cookies.Add($cookies)
# Scrape Microsoft's landing page for download link
$dlhtml = (Invoke-WebRequest $dlurl -WebSession $websession -UseBasicParsing) # get page with URL for exe # -UseBasicParsing # UseBasicParsing to supress cookie warning
$dllinks = ($dlhtml.Links | ForEach-Object { $_.href | Select-String "officedeployment"})
if (($dllinks.Length -gt 1) -and ($dllinks.Length -lt 100)){
  $dlexeurl = $dllinks[$dllinks.Length - 1].ToString().Trim()
} else {
  $dlexeurl = $dllinks.ToString().Trim()
}
Invoke-WebRequest $dlexeurl -WebSession $websession -UseBasicParsing -OutFile officedeploymenttool.exe # download officedeploymenttool.exe
## Extract setup from deployment tool
Write-Host "Extracting setup.exe..."
.\officedeploymenttool.exe /quiet /extract:$PWD
Start-Sleep -Seconds 5
## Build custom configuration file for unattended Office Pro Plus install
Write-Host "Building configuration.xml..."
Write-Output "" > configuration.xml
{<Configuration>
    <Add SourcePath="" OfficeClientEdition="32" Channel="Monthly">
      <Product ID="O365ProPlusRetail">
        <Language ID="en-us" />
        <ExcludeApp ID="InfoPath" />
        <ExcludeApp ID="SharePointDesigner" />
        <ExcludeApp ID="Groove" />
        <ExcludeApp ID="Lync" />
        <ExcludeApp ID="skypeforbusiness" />
        <ExcludeApp ID="Access" />
      </Product>
    </Add>
    <Display Level="None" AcceptEULA="TRUE" />
    <Updates Enabled="TRUE" Channel="Monthly" />
    <Property Name="AUTOACTIVATE" Value="1" />
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
    <Property Name="SharedComputerLicensing" Value="1" />
</Configuration>} > configuration.xml
Write-Host "Downloading deployment files using configuration.xml..."
.\setup.exe /download configuration.xml
## Configure setup with custom configuration file
Write-Host "Configuring using configuration.xml..."
Start-Sleep -Seconds 5
.\setup.exe /configure configuration.xml
## Create installation folder and set permissions on install files
Write-Host "Creating Install folder..."
mkdir OfficeInstaller
Write-Host "Fixing permissions on Office folder..."
takeown /F Office /R
$folders = Get-ChildItem Office -Recurse
Foreach ($folder in $folders) {
    ICACLS ("$($folder.FullName)") /grant ($env:USERNAME + ':(OI)(CI)F') /T
    Set-ItemProperty -LiteralPath $($folder.FullName) -Name IsReadOnly -Value $false
}
## Move installation files to installation folder
Write-Host "Moving install files to Install folder..."
Move-Item -Path Office -Destination OfficeInstaller\.
Move-Item -Path configuration.xml -Destination OfficeInstaller\.
Move-Item -Path setup.exe -Destination OfficeInstaller\.
## Create custom batch file for unattended Office 365 Pro Plus install
Write-Host "Creating Install365 batch file..."
Out-File -InputObject ('@ECHO OFF
TITLE Installing Office 365
ECHO Installing Office 365 (ProPlus)...
PUSHD %~dp0
CD /D %~dp0
SET "configFile=%~dp0configuration.xml"
SET "localDir=%~dp0"
powershell -Command "&{ $file =  Get-Item $env:configFile ; $xml = [xml](Get-Content $file) ; $xml.Configuration.Add.SourcePath = [string]$env:localDir ; $xml.Save($file.FullName) }"
%~dp0setup.exe /configure %~dp0configuration.xml
') -FilePath OfficeInstaller\Install365.bat -Encoding ascii -Force
Write-Host "Done creating Office 365 deployment package."
Write-Host "Use Install365.bat to install Microsoft Office 365"
