@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ========================================
echo   Date Folder Creator
echo ========================================
echo.

rem Get dropped folder path
if "%~1"=="" (
    echo Error: Please drag and drop a folder onto this batch file.
    echo.
    echo Usage:
    echo 1. Create a folder named like "2025.9-2026.5"
    echo 2. Drag and drop that folder onto this batch file
    echo 3. Date folders and subfolders will be created automatically
    echo.
    pause
    exit /b
)

set "TARGET_FOLDER=%~1"
set "FOLDER_NAME=%~nx1"

echo Dropped folder: %FOLDER_NAME%
echo.

rem Split period by hyphen
for /f "tokens=1,2 delims=-" %%a in ("%FOLDER_NAME%") do (
    set "START_DATE=%%a"
    set "END_DATE=%%b"
)

rem Check if parsing succeeded
if "%START_DATE%"=="" (
    echo Error: Folder name format is incorrect.
    echo Correct format: 2025.9-2026.5 or 2025.11-2026.3
    echo.
    pause
    exit /b
)

if "%END_DATE%"=="" (
    echo Error: Folder name format is incorrect.
    echo Correct format: 2025.9-2026.5 or 2025.11-2026.3
    echo.
    pause
    exit /b
)

rem Split start year and month
for /f "tokens=1,2 delims=." %%a in ("%START_DATE%") do (
    set "START_YEAR=%%a"
    set "START_MONTH_STR=%%b"
)

rem Split end year and month
for /f "tokens=1,2 delims=." %%a in ("%END_DATE%") do (
    set "END_YEAR=%%a"
    set "END_MONTH_STR=%%b"
)

rem Convert to numbers for comparison (remove leading zeros)
if "%START_MONTH_STR:~0,1%"=="0" (
    set /a START_MONTH=%START_MONTH_STR:~1%
) else (
    set /a START_MONTH=%START_MONTH_STR%
)

if "%END_MONTH_STR:~0,1%"=="0" (
    set /a END_MONTH=%END_MONTH_STR:~1%
) else (
    set /a END_MONTH=%END_MONTH_STR%
)

echo Period: %START_YEAR%/%START_MONTH_STR% - %END_YEAR%/%END_MONTH_STR%
echo.
echo Creating folders...
echo.

set /a FOLDER_COUNT=0

rem Loop through year and month
set /a CUR_YEAR=%START_YEAR%
set /a CUR_MONTH=%START_MONTH%
set "CUR_MONTH_STR=%START_MONTH_STR%"

:MONTH_LOOP
rem Check end condition
if %CUR_YEAR% GTR %END_YEAR% goto END_LOOP
if %CUR_YEAR% EQU %END_YEAR% if %CUR_MONTH% GTR %END_MONTH% goto END_LOOP

rem Get days in current month
for /f %%d in ('powershell -Command "[DateTime]::DaysInMonth(%CUR_YEAR%, %CUR_MONTH%)"') do set DAYS_IN_MONTH=%%d

echo Creating %CUR_YEAR%/%CUR_MONTH_STR% (%DAYS_IN_MONTH% days)...

rem Create folder for each day
for /l %%d in (1,1,%DAYS_IN_MONTH%) do (
    set "DATE_FOLDER=%CUR_YEAR%.%CUR_MONTH_STR%.%%d"
    
    rem Create date folder
    if not exist "%TARGET_FOLDER%\!DATE_FOLDER!" (
        mkdir "%TARGET_FOLDER%\!DATE_FOLDER!"
    )
    
    rem Create subfolder
    if not exist "%TARGET_FOLDER%\!DATE_FOLDER!\済" (
        mkdir "%TARGET_FOLDER%\!DATE_FOLDER!\済"
    )
    
    set /a FOLDER_COUNT+=1
)

rem Move to next month
set /a CUR_MONTH+=1
if %CUR_MONTH% GTR 12 (
    set /a CUR_MONTH=1
    set /a CUR_YEAR+=1
)

rem Update month string format
if %CUR_MONTH% LSS 10 (
    if "%START_MONTH_STR:~0,1%"=="0" (
        set "CUR_MONTH_STR=0%CUR_MONTH%"
    ) else (
        set "CUR_MONTH_STR=%CUR_MONTH%"
    )
) else (
    set "CUR_MONTH_STR=%CUR_MONTH%"
)

goto MONTH_LOOP

:END_LOOP

echo.
echo ========================================
echo   Completed!
echo ========================================
echo.
echo Created folders: %FOLDER_COUNT%
echo Subfolders created in each folder.
echo.
echo Location: %TARGET_FOLDER%
echo.
pause