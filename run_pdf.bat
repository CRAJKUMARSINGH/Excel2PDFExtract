@echo off
cd /d "%~dp0"
echo ================================================
echo  PDFExcelExtract - Budget PDF Generator
echo ================================================
echo.

REM Check Python exists
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo *** Python not found! ***
    echo.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to tick "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

python --version
echo.

REM Ensure required packages are installed
echo Installing/checking required packages...
python -m pip install --quiet openpyxl reportlab "xlrd==1.2.0"
echo Packages ready.
echo.

REM Run the script
echo Running apply.py ...
echo.
python apply.py
set EXIT=%ERRORLEVEL%

echo.
if %EXIT% EQU 0 (
    echo ================================================
    echo  SUCCESS: OUTPUT\budget_paged.pdf created!
    echo ================================================
    explorer "%~dp0OUTPUT"
) else (
    echo *** FAILED with exit code %EXIT% ***
)

echo.
pause
