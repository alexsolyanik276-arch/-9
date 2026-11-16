@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "QUARTAL_OPEN_BROWSER=1"
cd /d "%~dp0"
title Kvartal - Local App

echo ==============================================================
echo  KVARTAL - LOCAL APP
echo ==============================================================
echo.

if not exist "server.py" (
    echo ERROR: server.py was not found.
    echo Extract the complete project folder before running this file.
    echo.
    pause
    exit /b 1
)

set "PYTHON_CMD="
where py >nul 2>&1
if not errorlevel 1 (
    py -3 --version >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=py"
)

if not defined PYTHON_CMD (
    where python >nul 2>&1
    if not errorlevel 1 (
        python --version >nul 2>&1
        if not errorlevel 1 set "PYTHON_CMD=python"
    )
)

if not defined PYTHON_CMD (
    echo ERROR: Python 3 was not found.
    echo Install Python 3.10 or newer:
    echo https://www.python.org/downloads/windows/
    echo During setup, enable: Add python.exe to PATH
    echo Then run run_windows.bat again.
    echo.
    pause
    exit /b 1
)

echo Python was found. Starting the web server...
echo The browser should open automatically.
echo Keep this window open while using the application.
echo.

if "%PYTHON_CMD%"=="py" (
    py -3 server.py
) else (
    python server.py
)

set "EXIT_CODE=%ERRORLEVEL%"
echo.
if not "%EXIT_CODE%"=="0" (
    echo The application stopped with an error. Read the details above.
) else (
    echo The application has stopped.
)
echo.
echo Press any key to close this window.
pause >nul
exit /b %EXIT_CODE%
