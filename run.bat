@echo off
setlocal EnableDelayedExpansion
set "ROOT=%~dp0"
set "ROOT=%ROOT:~0,-1%"
set "PYTHON_EXE=%ROOT%\_runtime\python\python.exe"
if not exist "%PYTHON_EXE%" (echo [ERROR] python.exe not found at: %PYTHON_EXE% & echo Re-download the runtime. & pause & exit /b 1)
set "PLAYWRIGHT_BROWSERS_PATH=%ROOT%\_runtime\chromium"
set "PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1"
set "PYTHONPATH=%ROOT%\_runtime\python\libs"
echo SOP Generator v12 — پیشگامان پاکران الکترونیک
echo.
echo   Steps:    %ROOT%\edit\steps.xlsx
echo   Images:   %ROOT%\images
echo   Output:   %ROOT%\output\sop.pdf
echo.
echo   Generating PDF...
echo.
"%PYTHON_EXE%" "%ROOT%\app\generate_sop.py" --steps "%ROOT%\edit\steps.xlsx" --images "%ROOT%\images" --output "%ROOT%\output\sop.pdf" --template-dir "%ROOT%\app"
if errorlevel 1 (echo. & echo   ERROR: PDF generation failed. & pause & exit /b 1)
echo.
echo   PDF created: %ROOT%\output\sop.pdf
explorer /select,"%ROOT%\output\sop.pdf"
pause
