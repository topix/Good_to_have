========
@echo off
setlocal
rem set f=%0%
set f=SCVPN
title %f% - ® OpenSolution Nordic AB %date:~,4%
echo.
echo Stoppar OpenVPN service...
echo.
echo %f%
echo.
sc stop openvpnservice
echo.
echo Väntar några sekunder...
ping -n 3 127.0.0.1 >nul
echo.
echo Startar OpenVPN service...
echo.
sc config openvpnservice start= auto
sc start openvpnservice
echo.
echo Klar! (Stänger automatiskt om 10 sekunder)
ping -n 11 127.0.0.1 >nul