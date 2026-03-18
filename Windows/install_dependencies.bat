@echo off
echo ========================================
echo Employee Timesheet Processor
echo Installing Dependencies...
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo Python found! Installing dependencies...
echo.

REM Install dependencies
pip install pandas openpyxl

if errorlevel 1 (
    echo.
    echo ERROR: Installation failed
    echo Try running this file as Administrator
    pause
    exit /b 1
)

echo.
echo ========================================
echo SUCCESS! All dependencies installed
echo ========================================
echo.
echo You can now run the script with:
echo python process_timesheets.py your_file.xlsx
echo.
pause
