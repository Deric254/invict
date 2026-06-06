@echo off
:: Downloads and sets up embedded Python 3.12 into the .\python\ folder.
:: Called automatically by start.bat and BUILD.bat on first run.

set PYTHON_VERSION=3.12.4
set EMBED_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-embed-amd64.zip
set GET_PIP=https://bootstrap.pypa.io/get-pip.py
set DEST=%~dp0..\python

if exist "%DEST%\python.exe" (
    echo [OK] Embedded Python already exists.
    goto :install_packages
)

echo [1/4] Downloading Python %PYTHON_VERSION% embeddable...
curl -L --progress-bar -o "%TEMP%\py-embed.zip" "%EMBED_URL%"
if errorlevel 1 ( echo [ERROR] Download failed. Check internet connection. & exit /b 1 )

echo [2/4] Extracting...
mkdir "%DEST%" 2>nul
powershell -NoProfile -Command "Expand-Archive -Path '%TEMP%\py-embed.zip' -DestinationPath '%DEST%' -Force"
del "%TEMP%\py-embed.zip" 2>nul

echo [3/4] Enabling pip in embedded Python...
:: Un-comment 'import site' in the ._pth file so pip works
for %%F in ("%DEST%\python*._pth") do (
    powershell -NoProfile -Command ^
        "(Get-Content '%%F') -replace '#import site','import site' | Set-Content '%%F'"
)

echo [4/4] Installing pip...
curl -L --silent -o "%TEMP%\get-pip.py" "%GET_PIP%"
"%DEST%\python.exe" "%TEMP%\get-pip.py" --no-warn-script-location -q
del "%TEMP%\get-pip.py" 2>nul

:install_packages
echo [+] Installing required packages (openpyxl, reportlab)...
"%DEST%\python.exe" -m pip install openpyxl reportlab --quiet --no-warn-script-location
if errorlevel 1 ( echo [ERROR] Package install failed. Check internet. & exit /b 1 )

echo [OK] Python ready at: %DEST%
exit /b 0
