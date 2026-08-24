@echo off
setlocal EnableExtensions
rem Resolve a machine-local Python for VS Code/Cursor tasks.
rem Prefer the project venv, then E:\Programming\venv_%COMPUTERNAME%.
rem Never use experimental python*t.exe (win32com crashes).

set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"
set "PYEXE="
set "DEPTH=0"

if exist "%ROOT%\venv3\Scripts\python.exe" set "PYEXE=%ROOT%\venv3\Scripts\python.exe"
if not defined PYEXE if exist "%ROOT%\venv\Scripts\python.exe" set "PYEXE=%ROOT%\venv\Scripts\python.exe"
if not defined PYEXE if exist "%ROOT%\.venv\Scripts\python.exe" set "PYEXE=%ROOT%\.venv\Scripts\python.exe"
if defined PYEXE goto :Exec

set "SEARCH=%ROOT%"
:WalkHostVenv
if exist "%SEARCH%\venv_%COMPUTERNAME%\Scripts\python.exe" (
  set "PYEXE=%SEARCH%\venv_%COMPUTERNAME%\Scripts\python.exe"
  goto :Exec
)
for %%I in ("%SEARCH%\..") do set "PARENT=%%~fI"
if /I "%PARENT%"=="%SEARCH%" goto :CursorPython
set "SEARCH=%PARENT%"
set /a DEPTH+=1
if %DEPTH% GEQ 5 goto :CursorPython
goto :WalkHostVenv

:CursorPython
if not defined CURSOR_PYTHON goto :Appz
for %%I in ("%CURSOR_PYTHON%") do set "PYNAME=%%~nxI"
echo %PYNAME%| findstr /I /C:"t.exe" >nul
if %ERRORLEVEL%==0 (
  echo [task_python] skipping free-threaded interpreter: %CURSOR_PYTHON%
  goto :Appz
)
if exist "%CURSOR_PYTHON%" set "PYEXE=%CURSOR_PYTHON%"
if defined PYEXE goto :Exec

:Appz
if exist "C:\Appz\Python3.13\python.exe" set "PYEXE=C:\Appz\Python3.13\python.exe"
if defined PYEXE goto :Exec

where py >nul 2>&1
if %ERRORLEVEL%==0 (
  echo [task_python] using py -3
  py -3 %*
  exit /b %ERRORLEVEL%
)

echo [task_python] No usable Python found on %COMPUTERNAME%.
echo   Expected one of:
echo     %ROOT%\venv3\Scripts\python.exe
echo     %ROOT%\venv\Scripts\python.exe
echo     parent\venv_%COMPUTERNAME%\Scripts\python.exe
echo     Cursor-selected interpreter ^(not python*t.exe^)
echo     C:\Appz\Python3.13\python.exe
exit /b 1

:Exec
echo [task_python] using %PYEXE%
"%PYEXE%" %*
exit /b %ERRORLEVEL%
