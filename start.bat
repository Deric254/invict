@echo off
title St. Anne ICT Command Centre
cd /d "%~dp0"
echo ================================================
echo  St. Anne Mission Hospital -- ICT Command Centre
echo  Version 1.2 -- Excel-powered desktop app
echo ================================================
echo.

REM Prefer Python 3.11 or 3.12 — most stable with pywebview on Windows.
REM Python 3.13+ has a known pywebview accessibility recursion bug.
set PYTHON_CMD=

py -3.12 --version >nul 2>&1 && set PYTHON_CMD=py -3.12
if "%PYTHON_CMD%"=="" py -3.11 --version >nul 2>&1 && set PYTHON_CMD=py -3.11
if "%PYTHON_CMD%"=="" py -3.10 --version >nul 2>&1 && set PYTHON_CMD=py -3.10
if "%PYTHON_CMD%"=="" python --version >nul 2>&1 && set PYTHON_CMD=python
if "%PYTHON_CMD%"=="" python3 --version >nul 2>&1 && set PYTHON_CMD=python3
if "%PYTHON_CMD%"=="" py --version >nul 2>&1 && set PYTHON_CMD=py

if "%PYTHON_CMD%"=="" (
    echo ERROR: Python not found.
    echo.
    echo Please install Python 3.11 or 3.12 from:
    echo https://www.python.org/downloads/
    echo.
    echo IMPORTANT: During install, check "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

echo Python found: %PYTHON_CMD%
echo Installing required libraries...

REM Pin pywebview to 4.4.1 -- stable on Windows with EdgeChromium backend.
REM Later versions have a UI Automation recursion crash on Python 3.13+.
%PYTHON_CMD% -m pip install "pywebview==4.4.1" openpyxl --quiet --disable-pip-version-check

echo.
echo Starting ICT Command Centre...
echo.
%PYTHON_CMD% main.py

if errorlevel 1 (
    echo.
    echo === Application exited with an error ===
    echo If the window was blank, try running as Administrator.
    pause
)
