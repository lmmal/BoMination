@echo off
REM BoMination Application Deployment Script
REM This script helps deploy the BoMination application

echo ========================================
echo BoMination Application Deployment
echo ========================================
echo.

echo Checking for BoMinationApp.exe...
if exist "BoMinationApp.exe" (
    echo ✓ BoMinationApp.exe found
    echo.
    echo File Details:
    dir "BoMinationApp.exe" | findstr "BoMinationApp.exe"
    echo.
    echo Ready for deployment!
    echo.
    echo Instructions:
    echo 1. Copy BoMinationApp.exe to target computers
    echo 2. No installation required - just double-click to run
    echo 3. Refer to USER_GUIDE.md for operation instructions
    echo.
    echo Would you like to test the application now? (Y/N)
    set /p test_choice=
    if /i "%test_choice%"=="Y" (
        echo.
        echo Launching BoMination Application...
        start "BoMination" "BoMinationApp.exe"
    )
) else (
    echo ✗ BoMinationApp.exe not found!
    echo Please ensure you're running this script from the dist/ folder
    echo or copy BoMinationApp.exe to the current directory.
)

echo.
echo Deployment script completed.
pause
