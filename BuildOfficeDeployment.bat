@ECHO OFF
TITLE Building Office 365 deployment package
CD /D "%~dp0"
ECHO Downloading and building install package
powershell -ExecutionPolicy RemoteSigned -Command .\Builders\ScriptedDeploymentPrep.ps1

TITLE Building Office 365 compressed package installer
ECHO Packaging install package for distribution
CD /D "%~dp0"
powershell -ExecutionPolicy RemoteSigned -Command .\Builders\PackageOfficeInstaller.ps1