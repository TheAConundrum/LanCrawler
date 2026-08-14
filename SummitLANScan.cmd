@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title SummitLANScan

set "PY="
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY set "PY=python"

echo.
echo SummitLANScan — crawling all mapped network drives (no picker)
echo.

"%PY%" "%~dp0src\crawler\summit_lan_scan.py" %*
set "ERR=%ERRORLEVEL%"
echo.
pause
exit /b %ERR%
