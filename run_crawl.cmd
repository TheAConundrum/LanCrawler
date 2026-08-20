@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
title LAN Search Tool

set "ROOT=%~dp0"
set "PORTABLE_VENV=%ROOT%.portable_venv"
set "DEV_VENV=%ROOT%.venv"
set "WHEELS=%ROOT%vendor\wheels"
set "REQ=%ROOT%requirements-crawler.txt"
set "USED_PORTABLE=0"
set "PY="

if exist "%DEV_VENV%\Scripts\python.exe" (
  set "PY=%DEV_VENV%\Scripts\python.exe"
  goto :run
)

set "BOOTSTRAP="
where py >nul 2>&1
if not errorlevel 1 set "BOOTSTRAP=py -3"
if not defined BOOTSTRAP (
  where python >nul 2>&1
  if not errorlevel 1 set "BOOTSTRAP=python"
)
if not defined BOOTSTRAP (
  echo Python 3.10+ was not found.
  echo Install Python from python.org and check "Add python.exe to PATH", or use the py launcher.
  pause
  exit /b 1
)

if not exist "%WHEELS%\" (
  echo Missing vendor\wheels.
  echo Pack this kit with pack_portable.cmd on a machine that can pip download ^(JFrog^).
  echo This script will not install packages from the internet.
  pause
  exit /b 1
)

if not exist "%PORTABLE_VENV%\Scripts\python.exe" (
  echo Creating local Python packages folder...
  %BOOTSTRAP% -m venv "%PORTABLE_VENV%"
  if errorlevel 1 (
    echo Failed to create .portable_venv
    pause
    exit /b 1
  )
  "%PORTABLE_VENV%\Scripts\python.exe" -m pip install --no-index --find-links "%WHEELS%" --only-binary=:all: -r "%REQ%"
  if errorlevel 1 (
    echo.
    echo Failed to install packages from vendor\wheels.
    echo This PC Python:
    %BOOTSTRAP% -c "import sys,struct; print(sys.version); print(struct.calcsize('P')*8, 'bit')"
    echo.
    echo Packed wheels:
    dir /b "%WHEELS%\*.whl"
    echo.
    echo The wheel name must match this Python ^(cp310 / cp311 / cp312 / cp313 / cp314^) and 32/64-bit ^(win32 / win_amd64^).
    echo See ReadMe.txt. This script does not download from the internet.
    if exist "%PORTABLE_VENV%\" rmdir /s /q "%PORTABLE_VENV%"
    pause
    exit /b 1
  )
)

set "PY=%PORTABLE_VENV%\Scripts\python.exe"
set "USED_PORTABLE=1"

:run
"%PY%" "%ROOT%src\crawler\run_crawl.py" %*
set "ERR=%ERRORLEVEL%"

if "%USED_PORTABLE%"=="1" (
  echo.
  set /p DELPKGS=Do you want to delete the temporary python packages? [Y/N] 
  if /i "!DELPKGS!"=="Y" (
    echo Removing .portable_venv ...
    rmdir /s /q "%PORTABLE_VENV%"
  )
  echo.
  pause
)

if not "%USED_PORTABLE%"=="1" if not "%ERR%"=="0" pause
exit /b %ERR%
