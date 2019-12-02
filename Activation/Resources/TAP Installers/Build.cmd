@echo off
cls
pushd %~dp0

echo Building TAP Adapter (OpenVPN) Installer
cd InstallTAPAdapter
..\7z a -mx=9 "..\InstallTAPAdapter.7z" * 
cd ..
copy /b 7zsd.sfx + Config.txt + InstallTAPAdapter.7z ..\InstallTAPAdapter.exe

REM echo Building TAP Adapter (Stegnaos) Installer
REM cd InstallTAPAdapterOAS
REM ..\7z a -mx=9 "..\InstallTAPAdapterOAS.7z" * 
REM cd ..
REM copy /b 7zsd.sfx + Config.txt + InstallTAPAdapterOAS.7z ..\InstallTAPAdapterOAS.exe

echo Building TAP Adapter (Viscosity) Installer
cd InstallTAPAdapterViscosity
..\7z a -mx=9 "..\InstallTAPAdapterViscosity.7z" * 
cd ..
copy /b 7zsd.sfx + Config.txt + InstallTAPAdapterViscosity.7z ..\InstallTAPAdapterViscosity.exe

echo Cleaning Up
del *.7z
popd
