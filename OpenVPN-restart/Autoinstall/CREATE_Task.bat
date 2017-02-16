@echo off
setlocal
cd /d %~dp0

if not exist "C:\Utils\script_VPNRestart" mkdir C:\Utils\script_VPNRestart
COPY restartvpn.bat "C:\Utils\script_VPNRestart\restartvpn.bat"

REM Add Scheduled task
@echo off
schtasks /query > qvery
findstr /B /I "Restart OpenVPN" qvery >nul
if %errorlevel%==0  goto :delete
goto :create

:delete
SCHTASKS /DELETE /TN "Restart OpenVPN" /F >nul

:create
SCHTASKS /create /IT /RU "%USERNAME%" /SC DAILY /st 06:00:00 /TR "C:\Utils\script_VPNRestart\restartvpn.bat" /TN "Restart OpenVPN" >nul
del qvery >nul

REM Call Restart script
CALL "C:\Utils\script_VPNRestart\restartvpn.bat"
