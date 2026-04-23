@echo off
setlocal EnableExtensions
title Crimson Desert Teleporter Tool - Build EXE
cd /d "%~dp0"

set "SCRIPT_DIR=%~dp0"
set "LOCAL_PY=%SCRIPT_DIR%python\python.exe"
set "PY_LIBS=%SCRIPT_DIR%pylibs"
set "PYTHON="
set "PYTHON_SOURCE="
set "APP_NAME=CrimsonDesertTeleporter"

call :try_python "%LOCAL_PY%" "bundled Python"
if not defined PYTHON call :try_py_launcher
if not defined PYTHON call :try_where_python
if not defined PYTHON call :try_common_python_paths

if not defined PYTHON (
    echo No usable Python with tkinter was found.
    echo Run [run_teleporter.bat] once or install Python, then try again.
    pause
    exit /b 1
)

echo Using Python: "%PYTHON%"
if defined PYTHON_SOURCE echo Detected from: %PYTHON_SOURCE%
echo.

set "PYTHONPATH=%PY_LIBS%;%PYTHONPATH%"

set "NEED_INSTALL=0"
"%PYTHON%" -c "import pymem" >nul 2>&1
if errorlevel 1 set "NEED_INSTALL=1"
"%PYTHON%" -c "import webview" >nul 2>&1
if errorlevel 1 set "NEED_INSTALL=1"

if "%NEED_INSTALL%"=="1" (
    echo Missing dependencies. Installing to local folder...
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
        echo Failed to install dependencies for the build.
        pause
        exit /b 1
    )
)

"%PYTHON%" -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not found. Installing...
    "%PYTHON%" -m pip install pyinstaller
    if errorlevel 1 (
        echo Failed to install PyInstaller.
        pause
        exit /b 1
    )
)

echo Building %APP_NAME%.exe ...
"%PYTHON%" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --uac-admin ^
    --name "%APP_NAME%" ^
    --icon "%SCRIPT_DIR%teleporter.ico" ^
    --paths "%PY_LIBS%" ^
    --collect-submodules pymem ^
    --collect-submodules webview ^
    "%SCRIPT_DIR%cd_teleporter.py"

if errorlevel 1 (
    echo.
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Build complete:
echo   "%SCRIPT_DIR%dist\%APP_NAME%.exe"
pause
exit /b

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