@echo off
setlocal EnableExtensions
setlocal EnableDelayedExpansion
pushd "%~dp0"
reg.exe query "hklm\software\microsoft\Windows NT\currentversion" /v buildlabex | find /i "amd64" >nul 2>&1
if %errorlevel% equ 0 set xOS=x64
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set xOS=x86
certutil -addstore "TrustedPublisher" openvpn.cer >nul 2>&1
%xOS%\tapinstall.exe install %xOS%\OemVista.inf tap0901 >nul 2>&1
del /f /s /q %TEMP%\InstallTAPAdapter.exe
exit