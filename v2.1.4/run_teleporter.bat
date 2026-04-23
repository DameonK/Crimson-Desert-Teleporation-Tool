@echo off
setlocal EnableExtensions
title Crimson Desert Teleporter Tool - Launcher
cd /d "%~dp0"

:: Check for admin rights.
net session >nul 2>&1
if errorlevel 1 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process -FilePath '%~f0' -WorkingDirectory '%~dp0' -Verb RunAs"
    exit /b
)

set "SCRIPT_DIR=%~dp0"
set "LOCAL_PY=%SCRIPT_DIR%python\python.exe"
set "PY_DIR=%SCRIPT_DIR%python"
set "PY_LIBS=%SCRIPT_DIR%pylibs"
set "PYTHON="
set "PYTHON_SOURCE="

:: Prefer bundled Python. Otherwise find a real interpreter path and
:: avoid the Windows Store stub under WindowsApps.
call :try_python "%LOCAL_PY%" "bundled Python"
if not defined PYTHON call :try_py_launcher
if not defined PYTHON call :try_where_python
if not defined PYTHON call :try_common_python_paths
if not defined PYTHON goto :install_local_python

:found_python
echo Using Python: "%PYTHON%"
if defined PYTHON_SOURCE echo Detected from: %PYTHON_SOURCE%
echo.

set "PYTHONPATH=%PY_LIBS%;%PYTHONPATH%"

:: Install missing dependencies to local folder.
set "NEED_INSTALL=0"
"%PYTHON%" -c "import pymem" >nul 2>&1
if errorlevel 1 set "NEED_INSTALL=1"
"%PYTHON%" -c "import webview" >nul 2>&1
if errorlevel 1 set "NEED_INSTALL=1"

if "%NEED_INSTALL%"=="1" (
    echo Installing missing dependencies to local folder...
    if not exist "%PY_LIBS%" mkdir "%PY_LIBS%"

    "%PYTHON%" -m pip --version >nul 2>&1
    if errorlevel 1 (
        echo pip is missing. Bootstrapping pip...
        "%PYTHON%" -m ensurepip --upgrade >nul 2>&1
        if errorlevel 1 (
            echo Failed to enable pip for "%PYTHON%".
            pause
            exit /b 1
        )
    )

    "%PYTHON%" -m pip install pymem pywebview --target="%PY_LIBS%" --quiet --no-warn-script-location --no-deps
    "%PYTHON%" -m pip install bottle proxy_tools --target="%PY_LIBS%" --quiet --no-warn-script-location
    if errorlevel 1 (
        echo Failed to install dependencies.
        echo Python used: "%PYTHON%"
        pause
        exit /b 1
    )
    echo Done.
)

"%PYTHON%" "%SCRIPT_DIR%cd_teleporter.py"
if errorlevel 1 (
    echo.
    echo Program exited with an error.
    pause
)
exit /b

:install_local_python
echo Python with tkinter was not found. Downloading local Python (one-time setup)...
echo.

set "PY_VERSION=3.12.7"
set "PY_EXE=python-%PY_VERSION%-amd64.exe"
set "PY_URL=https://www.python.org/ftp/python/%PY_VERSION%/%PY_EXE%"

echo Downloading Python %PY_VERSION% installer (~30 MB)...
powershell -Command ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " ^
    "$ProgressPreference = 'SilentlyContinue'; " ^
    "Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%SCRIPT_DIR%%PY_EXE%'"
if not exist "%SCRIPT_DIR%%PY_EXE%" (
    echo.
    echo Failed to download Python. Check your internet connection.
    pause
    exit /b 1
)

echo Installing Python to local folder (no system changes)...
"%SCRIPT_DIR%%PY_EXE%" /quiet TargetDir="%PY_DIR%" ^
    InstallAllUsers=0 PrependPath=0 Include_launcher=0 ^
    Include_test=0 Include_doc=0 Include_tcltk=1 ^
    AssociateFiles=0 Shortcuts=0 Include_pip=1

del "%SCRIPT_DIR%%PY_EXE%" 2>nul

call :try_python "%LOCAL_PY%" "bundled local install"
if not defined PYTHON (
    echo.
    echo Failed to install Python. Try running the launcher again.
    pause
    exit /b 1
)

echo Python %PY_VERSION% installed to local folder.
echo.
goto :found_python

:try_py_launcher
for /f "usebackq delims=" %%P in (`py -3 -c "import sys, tkinter; print(sys.executable)" 2^>nul`) do (
    if not defined PYTHON set "PYTHON=%%P"
)
if defined PYTHON set "PYTHON_SOURCE=py launcher"
exit /b

:try_where_python
for /f "delims=" %%P in ('where python 2^>nul') do call :try_python "%%~fP" "PATH"
exit /b

:try_common_python_paths
for %%P in ("%LocalAppData%\Python\*\python.exe" "%LocalAppData%\Programs\Python\Python*\python.exe" "%ProgramFiles%\Python*\python.exe" "%ProgramFiles(x86)%\Python*\python.exe") do call :try_python "%%~fP" "common install path"
exit /b

:try_python
if defined PYTHON exit /b 0
set "CANDIDATE=%~1"
if not defined CANDIDATE exit /b 1
if not exist "%CANDIDATE%" exit /b 1

if /I not "%CANDIDATE:\WindowsApps\=%"=="%CANDIDATE%" exit /b 1

"%CANDIDATE%" -c "import tkinter" >nul 2>&1
if errorlevel 1 exit /b 1

for /f "usebackq delims=" %%P in (`"%CANDIDATE%" -c "import sys; print(sys.executable)" 2^>nul`) do (
    if not defined PYTHON set "PYTHON=%%P"
)
if defined PYTHON set "PYTHON_SOURCE=%~2"
exit /b