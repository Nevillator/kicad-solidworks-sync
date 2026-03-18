@echo off
:: Uninstall KiCad-SolidWorks Sync Add-in
:: Must be run as Administrator

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: This script must be run as Administrator.
    echo Right-click and select "Run as administrator".
    pause
    exit /b 1
)

set INSTALLED=%~dp0solidworks-addin\KiCadSync\bin\installed

echo Unregistering KiCad Sync add-in...
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" /unregister "%INSTALLED%\KiCadSync.dll"

if %errorlevel% equ 0 (
    echo.
    echo Unregistered. Restart SolidWorks to complete removal.
) else (
    echo.
    echo Unregistration failed. Check the errors above.
)

pause
