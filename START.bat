@echo off
chcp 65001 >nul
title CRM Export Tool

cd /d "%~dp0"

echo.
echo ========================================
echo   CRM Export and Auto-Call Tool
echo ========================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found!
    echo Please download and install Python from:
    echo https://www.python.org/downloads/
    echo.
    echo IMPORTANT: check "Add Python to PATH" during install.
    echo Then run this file again.
    echo.
    start https://www.python.org/downloads/
    pause
    exit
)

echo Installing/checking required libraries...
pip install selenium beautifulsoup4 webdriver-manager openpyxl requests -q
echo.

:menu
echo ========================================
echo   Choose an action:
echo ========================================
echo   1. Export orders from CRM (and prepare MTS file)
echo   2. Merge MTS call report with CRM orders
echo   3. Exit
echo ========================================
set /p choice="Enter 1, 2 or 3: "

if "%choice%"=="1" goto crm
if "%choice%"=="2" goto merge
if "%choice%"=="3" goto end
echo Invalid choice, try again.
goto menu

:crm
echo.
echo Starting CRM export...
echo.
python scrape_crm.py crm
goto finish

:merge
echo.
set "reportfile=report.csv"
set /p "reportfile=Enter report filename (press Enter for report.csv): "
echo.
echo Using file: %reportfile%
python scrape_crm.py merge "%reportfile%"
goto finish

:finish
echo.
echo ========================================
echo Done.
echo ========================================
pause
goto menu

:end
exit
