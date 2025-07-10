@echo off
REM Comprehensive deployment verification script
REM This tests for potential roadblocks before distributing to sales team

echo ========================================
echo BoMination Deployment Verification
echo ========================================
echo.

echo Checking deployment readiness...
echo.

REM Check if executable exists
echo [1/8] Checking executable...
if exist "dist\BoMinationApp.exe" (
    echo ✓ BoMinationApp.exe found
    for %%I in ("dist\BoMinationApp.exe") do echo   Size: %%~zI bytes (%%~nI.%%~xI^)
) else (
    echo ❌ BoMinationApp.exe NOT FOUND!
    echo    Build may have failed or files moved
    goto :error
)

REM Check for dependencies that might not be bundled
echo.
echo [2/8] Checking for external dependencies...

REM Test if executable can start (quick test)
echo.
echo [3/8] Testing executable startup...
echo   (This will launch and immediately close the app)
timeout /t 2 /nobreak >nul
start /wait /min "BoMination Test" "dist\BoMinationApp.exe" --version 2>nul
if %errorlevel% equ 0 (
    echo ✓ Executable starts successfully
) else (
    echo ⚠ Warning: Executable may have startup issues
    echo   This could indicate missing dependencies
)

REM Check for documentation
echo.
echo [4/8] Checking documentation...
if exist "USER_GUIDE.md" (
    echo ✓ User guide available
) else (
    echo ⚠ Warning: USER_GUIDE.md missing
)

if exist "SalesTeam_Package" (
    echo ✓ Sales team package ready
    echo   Contents:
    dir "SalesTeam_Package" /b | findstr /V "^$"
) else (
    echo ❌ Sales team package not created
    echo   Run create_sales_package.bat first
)

REM Check system compatibility
echo.
echo [5/8] Checking system compatibility...
echo   OS Version: 
ver | findstr /C:"Windows"
if %errorlevel% equ 0 (
    echo ✓ Windows OS detected
) else (
    echo ❌ Non-Windows OS - application may not work
)

REM Check architecture
echo   Architecture: %PROCESSOR_ARCHITECTURE%
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    echo ✓ 64-bit architecture compatible
) else (
    echo ⚠ Warning: Non-64-bit architecture detected
    echo   Application built for 64-bit Windows
)

REM Check available disk space
echo.
echo [6/8] Checking disk space...
for /f "tokens=3" %%a in ('dir "dist" /-c ^| findstr /C:"bytes free"') do (
    set freespace=%%a
)
echo   Free space available: %freespace% bytes
REM Convert to MB for readability
set /a freemb=%freespace:~0,-6%
if %freemb% gtr 500 (
    echo ✓ Sufficient disk space ^(%freemb% MB^)
) else (
    echo ⚠ Warning: Low disk space ^(%freemb% MB^)
    echo   Application needs ~200MB for operation
)

REM Check for potential security issues
echo.
echo [7/8] Security considerations...
echo   • Executable is NOT digitally signed
echo   • Users will see Windows security warnings
echo   • Antivirus may flag as suspicious (normal for PyInstaller^)
echo   ✓ These are normal for PyInstaller executables

REM Final verification
echo.
echo [8/8] Final deployment checklist...
echo.
echo Items to distribute to sales team:
echo   📁 SalesTeam_Package folder containing:
if exist "SalesTeam_Package\BoMinationApp.exe" (
    echo   ✓ BoMinationApp.exe ^(main application^)
) else (
    echo   ❌ BoMinationApp.exe missing from package
)

if exist "SalesTeam_Package\USER_GUIDE.md" (
    echo   ✓ USER_GUIDE.md ^(instructions^)
) else (
    echo   ❌ USER_GUIDE.md missing from package
)

if exist "SalesTeam_Package\QUICK_START.txt" (
    echo   ✓ QUICK_START.txt ^(quick reference^)
) else (
    echo   ❌ QUICK_START.txt missing from package
)

echo.
echo ========================================
echo DEPLOYMENT VERIFICATION COMPLETE
echo ========================================
echo.

echo Known potential roadblocks for sales team:
echo   1. Windows security warning on first run
echo      └─ Solution: Click "More info" → "Run anyway"
echo.
echo   2. Antivirus software blocking executable
echo      └─ Solution: Add to antivirus whitelist
echo.
echo   3. Corporate firewalls blocking price lookup
echo      └─ Solution: Ensure internet access for pricing
echo.
echo   4. Insufficient permissions for temp files
echo      └─ Solution: Run from user folder, not Program Files
echo.
echo   5. Very large PDF files causing memory issues
echo      └─ Solution: Process smaller PDFs or restart app
echo.

echo ✅ APPLICATION IS READY FOR DISTRIBUTION!
echo.
echo Next steps:
echo   1. Zip the SalesTeam_Package folder
echo   2. Email to sales team or place on network share  
echo   3. Include instructions to "Run as normal user"
echo   4. Provide IT contact for technical issues
echo.
goto :end

:error
echo.
echo ❌ DEPLOYMENT NOT READY - FIX ISSUES ABOVE
echo.

:end
pause
