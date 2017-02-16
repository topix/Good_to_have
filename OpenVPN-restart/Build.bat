@echo off
setlocal
cd /d %~dp0

copy /b 7zsd_All.sfx + 1-GLOBAL_autodestroy.txt + Autoinstall.7z Restart-OpenVPN-X.exe
copy /b 7zsd_All.sfx + 1-GLOBAL_standard.txt + Autoinstall.7z Restart-OpenVPN.exe
