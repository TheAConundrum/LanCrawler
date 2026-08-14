@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PY="
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY set "PY=python"

"%PY%" "%~dp0src\crawler\run_crawl.py" %*
exit /b %ERRORLEVEL%
