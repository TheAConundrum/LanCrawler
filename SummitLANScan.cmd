@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "LAN_CRAWL_HIDE_CONSOLE=1"

set "PY="
set "PYW="
if exist "%~dp0.venv\Scripts\pythonw.exe" set "PYW=%~dp0.venv\Scripts\pythonw.exe"
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY set "PY=python"
if not defined PYW (
  for %%I in ("%PY%") do if exist "%%~dpIpythonw.exe" set "PYW=%%~dpIpythonw.exe"
)
if not defined PYW (
  where pythonw >nul 2>&1
  if not errorlevel 1 set "PYW=pythonw"
)

if defined PYW (
  start "" "%PYW%" "%~dp0src\crawler\summit_lan_scan.py" %*
  exit /b 0
)

"%PY%" "%~dp0src\crawler\summit_lan_scan.py" %*
exit /b %ERRORLEVEL%
