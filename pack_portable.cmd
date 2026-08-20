@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Pack LAN Search Tool zip

set "PY="
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY (
  where py >nul 2>&1
  if not errorlevel 1 set "PY=py -3"
)
if not defined PY set "PY=python"

echo Packing dist\Lan_Search_Tool.zip (pip download uses this PC's JFrog / pip index)...
%PY% "%~dp0src\pack_portable.py"
set "ERR=%ERRORLEVEL%"
echo.
pause
exit /b %ERR%
