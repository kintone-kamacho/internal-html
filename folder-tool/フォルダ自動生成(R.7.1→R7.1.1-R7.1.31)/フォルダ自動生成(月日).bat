@echo off
setlocal enabledelayedexpansion

set "TARGET_FOLDER=%~1"

if "%TARGET_FOLDER%"=="" (
    echo Error: Please drag and drop a folder onto this batch file.
    pause
    exit /b 1
)

if not exist "%TARGET_FOLDER%" (
    echo Error: Specified folder does not exist.
    pause
    exit /b 1
)

for %%F in ("%TARGET_FOLDER%") do set "FOLDER_NAME=%%~nxF"
for /f "tokens=2 delims=." %%M in ("%FOLDER_NAME%") do set "MONTH=%%M"

if "%MONTH%"=="" (
    echo Error: Could not extract month from folder name.
    echo Folder name should be like "R7.12"
    pause
    exit /b 1
)

echo Creating folders in: %FOLDER_NAME%
echo Month: %MONTH%
echo.

for /L %%D in (1,1,31) do (
    set "DAY_FOLDER=%TARGET_FOLDER%\%MONTH%.%%D"
    if not exist "!DAY_FOLDER!" (
        mkdir "!DAY_FOLDER!"
        echo Created: %MONTH%.%%D
    ) else (
        echo Exists: %MONTH%.%%D
    )
)

echo.
echo Complete!
pause