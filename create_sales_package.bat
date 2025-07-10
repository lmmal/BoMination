@echo off
REM Create deployment package for Omni Sales Team
echo ========================================
echo Creating BoMination Sales Team Package
echo ========================================
echo.

REM Create a clean package directory
if exist "SalesTeam_Package" rmdir /s /q "SalesTeam_Package"
mkdir "SalesTeam_Package"

echo Copying files for sales team deployment...

REM Copy the essential files
copy "dist\BoMinationApp.exe" "SalesTeam_Package\" >nul
if exist "USER_GUIDE.md" copy "USER_GUIDE.md" "SalesTeam_Package\" >nul
copy "dist\deploy.bat" "SalesTeam_Package\" >nul

REM Create a simple README for the package
echo Creating Quick Start instructions...
(
echo # BoMination Application - Quick Start
echo.
echo ## For Sales Team Members:
echo 1. Double-click `BoMinationApp.exe` to launch the application
echo 2. If Windows shows a security warning, click "More info" then "Run anyway"
echo 3. Follow the USER_GUIDE.md for detailed instructions
echo.
echo ## What's Included:
echo - BoMinationApp.exe ^(90.3 MB^) - The main application
echo - USER_GUIDE.md - Step-by-step instructions
echo - deploy.bat - Optional test script
echo.
echo ## No Installation Required!
echo This is a portable application - just double-click and run.
echo.
echo For technical support, contact your IT department.
) > "SalesTeam_Package\QUICK_START.txt"

echo.
echo ✓ Package created in 'SalesTeam_Package' folder
echo.
echo Contents:
dir "SalesTeam_Package" /b
echo.

REM Calculate total size
for /f "tokens=3" %%a in ('dir "SalesTeam_Package" /s /-c ^| findstr /c:" File(s)"') do set totalsize=%%a
set /a totalmb=%totalsize:~0,-3%
echo Total package size: ~%totalmb% MB
echo.

echo Package ready for distribution!
echo.
echo Next steps:
echo 1. Test the package by running deploy.bat in SalesTeam_Package folder
echo 2. Zip the SalesTeam_Package folder for email distribution
echo 3. Or copy the folder contents to a shared network drive
echo.
pause
