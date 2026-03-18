@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ========================================
echo Employee Timesheet Processor
echo ========================================
echo.

REM Check if file was provided
if "%~1"=="" (
    echo ERROR: No file specified
    echo Drag and drop your Excel file onto this batch file
    echo.
    pause
    exit /b 1
)

REM Check if file exists
if not exist "%~1" (
    echo ERROR: File not found: %~1
    echo.
    pause
    exit /b 1
)

REM Get paths
set "INPUT=%~1"
set "SCRIPTDIR=%~dp0"
set "LOGFILE=%SCRIPTDIR%_process.log"

echo Processing: %~nx1
echo.

REM Run Python script and capture output
python "%SCRIPTDIR%process_timesheets.py" "%INPUT%" "%SCRIPTDIR%_temp.xlsx" > "%LOGFILE%" 2>&1
if errorlevel 1 (
    py "%SCRIPTDIR%process_timesheets.py" "%INPUT%" "%SCRIPTDIR%_temp.xlsx" > "%LOGFILE%" 2>&1
    if errorlevel 1 (
        echo ERROR: Processing failed
        echo.
        type "%LOGFILE%"
        del "%LOGFILE%" 2>nul
        pause
        exit /b 1
    )
)

REM Extract date from log
set "DATE="
for /f "tokens=4 delims= " %%a in ('findstr /C:"Filtering for date:" "%LOGFILE%"') do set "DATE=%%a"

REM Show log
type "%LOGFILE%"
echo.

REM Determine output filename
if defined DATE (
    set "OUTPUT=rekordok_%DATE%.xlsx"
) else (
    echo Warning: Could not extract date
    for /f %%a in ('powershell -Command "Get-Date -Format yyyy-MM-dd"') do set "OUTPUT=rekordok_%%a.xlsx"
)

REM Move temp file to final name
if exist "%SCRIPTDIR%_temp.xlsx" (
    move /Y "%SCRIPTDIR%_temp.xlsx" "%SCRIPTDIR%!OUTPUT!" >nul
    if exist "%SCRIPTDIR%!OUTPUT!" (
        echo ========================================
        echo SUCCESS: !OUTPUT!
        echo ========================================
    ) else (
        echo ERROR: Failed to create output file
    )
) else (
    echo ERROR: Temp file not created
)

REM Cleanup
del "%LOGFILE%" 2>nul

echo.
pause
