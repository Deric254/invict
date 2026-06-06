@echo off
title St. Anne ICT Command Centre — Installer Builder
color 0A
cd /d "%~dp0"

echo ============================================================
echo   St. Anne Mission Hospital — ICT Command Centre
echo   Building Installer EXE
echo ============================================================
echo.

:: ── Node.js check ────────────────────────────────────────────────────────────
node --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Node.js not found. Install from https://nodejs.org
    pause & exit /b 1
)
for /f "tokens=* delims=v" %%V in ('node --version') do set NODE_VER=%%V
echo [OK] Node.js v%NODE_VER%

:: ── Bundled Python ────────────────────────────────────────────────────────────
if not exist "python\python.exe" (
    echo.
    echo [INFO] Downloading embedded Python - needs internet, ~2 min...
    call scripts\bundle-python.bat
    if errorlevel 1 ( pause & exit /b 1 )
) else (
    echo [OK] Embedded Python ready.
)

:: ── Snapshot your live data into assets\ ─────────────────────────────────────
echo.
echo [INFO] Snapshotting your current data into the installer...

:: auth.json — carries your PIN into the installer
if exist "data-dev\auth.json" (
    copy /Y "data-dev\auth.json" "assets\auth.json" >nul
    echo [OK] PIN snapshot taken from data-dev\auth.json
) else if exist "auth.json" (
    copy /Y "auth.json" "assets\auth.json" >nul
    echo [OK] PIN snapshot taken from auth.json
) else (
    echo [INFO] No auth.json found - installer will prompt PIN on first run
)

:: ICT_MASTER.xlsx — carry live inventory into the installer
if exist "data-dev\ICT_MASTER.xlsx" (
    copy /Y "data-dev\ICT_MASTER.xlsx" "assets\ICT_MASTER.xlsx" >nul
    echo [OK] Inventory snapshot taken from data-dev\ICT_MASTER.xlsx
) else if exist "ICT_MASTER.xlsx" (
    copy /Y "ICT_MASTER.xlsx" "assets\ICT_MASTER.xlsx" >nul
    echo [OK] Inventory snapshot taken from ICT_MASTER.xlsx
) else (
    echo [INFO] No ICT_MASTER.xlsx found - installer will create fresh database
)

:: ── Node dependencies ─────────────────────────────────────────────────────────
if not exist "node_modules\electron" (
    echo.
    echo [INFO] Installing build tools - one time only...
    npm install
    if errorlevel 1 ( echo [ERROR] npm install failed. & pause & exit /b 1 )
)
echo [OK] Build tools ready.

:: ── Build installer ───────────────────────────────────────────────────────────
echo.
echo Building installer... (3-6 minutes, downloading Electron if first time)
echo.
npm run build:win
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed - see messages above.
    pause & exit /b 1
)

echo.
echo ============================================================
echo   DONE!
echo.
echo   dist\ICT_CommandCentre_Setup.exe
echo.
echo   This is a proper installer - send it to any Windows PC.
echo   It will install to Program Files with a Start Menu shortcut.
echo   Your PIN and data are bundled inside.
echo ============================================================
echo.
pause
