@echo off
REM ============================================================
REM  SOP Portable v12 — Runtime setup
REM
REM  Downloads:
REM    - Python 3.12 embeddable (zip)
REM    - pip bootstrap (get-pip.py)
REM    - Python wheels: openpyxl, Pillow, Jinja2, playwright, pypdf
REM    - Chromium browser (via playwright install)
REM
REM  Installs into:
REM    _runtime/python/         Python executable + stdlib
REM    _runtime/python/libs/    third-party wheels
REM    _runtime/chromium/       Chromium browser
REM
REM  Idempotent — re-running only fetches missing pieces.
REM ============================================================
setlocal enabledelayedexpansion

set "ROOT=%~dp0"
if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"
set "RT=%ROOT%\_runtime"
set "PY_DIR=%RT%\python"
set "LIBS_DIR=%RT%\python\libs"
set "PW_DIR=%RT%\chromium"
set "PY_EXE=%PY_DIR%\python.exe"

echo ============================================================
echo  SOP Portable v12 — Runtime setup
echo  ROOT:   %ROOT%
echo  Target: %RT%
echo ============================================================

REM --- Create dirs ---
if not exist "%RT%" mkdir "%RT%"
if not exist "%PY_DIR%" mkdir "%PY_DIR%"
if not exist "%LIBS_DIR%" mkdir "%LIBS_DIR%"
if not exist "%PW_DIR%" mkdir "%PW_DIR%"

set "CURL=curl -L --silent --show-error --fail"

REM --- Step 1: Python 3.12 embeddable ---
if exist "%PY_EXE%" (
    echo [1/4] Python embeddable: already present
) else (
    echo [1/4] Downloading Python 3.12 embeddable...
    set "PY_VER=3.12.7"
    set "PY_URL=https://www.python.org/ftp/python/!PY_VER!/python-!PY_VER!-embed-amd64.zip"
    set "PY_ZIP=%RT%\python-embed.zip"
    %CURL% -o "%PY_ZIP%" "!PY_URL!"
    if errorlevel 1 (
        echo ERROR: failed to download Python from !PY_URL!
        exit /b 1
    )
    echo  Extracting...
    powershell -NoProfile -Command "Expand-Archive -Path '%PY_ZIP%' -DestinationPath '%PY_DIR%' -Force"
    del "%PY_ZIP%"
    if not exist "%PY_EXE%" (
        echo ERROR: python.exe not found after extraction
        exit /b 1
    )
    REM Enable site-packages (uncomment _pth line)
    set "PTH=%PY_DIR%\python312._pth"
    if exist "!PTH!" (
        powershell -NoProfile -Command "(Get-Content '!PTH!') -replace '^#import site', 'import site' | Set-Content '!PTH!'"
        REM Append libs dir to sys.path
        echo Lib\libs>>"!PTH!"
        echo ..\libs>>"!PTH!"
    )
)

REM --- Step 2: pip bootstrap ---
"%PY_EXE%" -m pip --version >nul 2>&1
if not errorlevel 1 (
    echo [2/4] pip: already available
) else (
    echo [2/4] Bootstrapping pip...
    set "GETPIP=%RT%\get-pip.py"
    %CURL% -o "!GETPIP!" "https://bootstrap.pypa.io/get-pip.py"
    if errorlevel 1 (
        echo ERROR: failed to download get-pip.py
        exit /b 1
    )
    "%PY_EXE%" "!GETPIP!" --no-warn-script-location
    del "!GETPIP!"
)

REM --- Step 3: Wheels ---
echo [3/4] Installing wheels into %LIBS_DIR% ...
set "WHEELS=openpyxl Pillow Jinja2 playwright pypdf"
for %%w in (%WHEELS%) do (
    echo  - %%w
    "%PY_EXE%" -m pip install --upgrade --target="%LIBS_DIR%" --no-warn-script-location %%w
    if errorlevel 1 (
        echo ERROR: failed to install %%w
        exit /b 1
    )
)

REM --- Step 4: Chromium ---
set "CHROME_FOUND="
if exist "%PW_DIR%\chromium-*\chrome-win64\chrome.exe" set "CHROME_FOUND=1"
if exist "%PW_DIR%\chromium-*\chrome-win\chrome.exe" set "CHROME_FOUND=1"
if defined CHROME_FOUND (
    echo [4/4] Chromium: already present
) else (
    echo [4/4] Downloading Chromium via Playwright...
    set "PLAYWRIGHT_BROWSERS_PATH=%PW_DIR%"
    set "PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=0"
    "%PY_EXE%" -m playwright install chromium
    if errorlevel 1 (
        echo ERROR: playwright install chromium failed
        exit /b 1
    )
)

REM --- Final sanity check ---
echo.
echo ============================================================
echo  Runtime ready.
echo  Python:    %PY_EXE%
echo  Packages:  %LIBS_DIR%
echo  Chromium:  %PW_DIR%
echo ============================================================
echo  Next: run build_excels.py to create starter Excels, then run.bat
echo.

endlocal
