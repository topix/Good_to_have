@echo off
setlocal
cd /d %~dp0

REM -- TV_Host --
copy /b 7zsd_All.sfx + Build_Settings.txt + installer-host.7z TV_Host.exe
copy /b 7zsd_All.sfx + Build_Settings.txt + installer-host-ar.7z TV_Host-AR.exe

REM AR stands for auto restart
