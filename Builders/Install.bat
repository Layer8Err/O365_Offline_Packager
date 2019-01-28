@ECHO OFF
TITLE Decompress Microsoft Office
ECHO Decompressing Microsoft Office...
cd /D "%~dp0"
powershell -ExecutionPolicy RemoteSigned -Command .\Decompress_install_Office.ps1
