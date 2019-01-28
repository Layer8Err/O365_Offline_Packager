# O365_Offline_Packager
Create distributable offline Office 365 ProPlus install package.

This script will automatically download the latest Office Deployment Tool from Microsoft
and build an OfficeInstaller for deployment.
Once complete, the installation will be placed in the "OfficeInstaller" folder with an 
install script and configuration file.
The "OfficeInstaller" folder will then be compressed and added to a new folder called
"Office_365_CompressedInstall" that will contain a batch and PowerShell installer to
decompress and install Office. You can manually update this install package by replacing
the compressed "CompressedOfficeInstaller.zip" with a newer zip file.

By default, this is configured for Microsoft Office 365 Pro Plus 32-bit


## Manual setup:

Get [Office Deployment Tool from Microsoft](https://www.microsoft.com/en-us/download/details.aspx?id=49117)

* NOTE: Executible download link will change as Microsoft updates the deployment tool

1. Extract deployment tool "setup.exe" file:

``` .\officedeploymenttool_*.exe /extract:$pwd ```

2. Download deployment files:

``` .\setup.exe /download```

3. Configure using configuration.xml

``` .\setup.exe /configure configuration.xml```