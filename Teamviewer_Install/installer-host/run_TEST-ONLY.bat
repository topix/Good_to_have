@echo off
setlocal
cd /d %~dp0

REM Remove old 
cscript uninst-inst.vbs
