@echo off
:: Install KiCad-SolidWorks Sync Add-in
:: Must be run as Administrator
:: SolidWorks must be CLOSED

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: This script must be run as Administrator.
    echo Right-click and select "Run as administrator".
    pause
    exit /b 1
)

set STAGING=%~dp0solidworks-addin\KiCadSync\bin\staging\net48
set INSTALLED=%~dp0solidworks-addin\KiCadSync\bin\installed

if not exist "%STAGING%\KiCadSync.dll" (
    echo ERROR: No staged build found at:
    echo   %STAGING%
    echo.
    echo Build the project first.
    pause
    exit /b 1
)

echo Copying build to installed location...
if not exist "%INSTALLED%" mkdir "%INSTALLED%"
xcopy /Y /Q "%STAGING%\*.*" "%INSTALLED%\"

echo Registering KiCad Sync add-in with SolidWorks...
"%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" /codebase "%INSTALLED%\KiCadSync.dll"

if %errorlevel% equ 0 (
    echo.
    echo Success! Start SolidWorks to load the add-in.
) else (
    echo.
    echo Registration failed. Check the errors above.
)

pause
