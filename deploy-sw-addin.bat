@echo off
:: Deploy latest build to the installed location
:: SolidWorks must be CLOSED (the DLL is locked while SW is running)

set STAGING=%~dp0solidworks-addin\KiCadSync\bin\staging\net48
set INSTALLED=%~dp0solidworks-addin\KiCadSync\bin\installed

if not exist "%STAGING%\KiCadSync.dll" (
    echo ERROR: No staged build found. Run build first.
    pause
    exit /b 1
)

xcopy /Y /Q "%STAGING%\*.*" "%INSTALLED%\"
if %errorlevel% equ 0 (
    echo Deployed. Start SolidWorks.
) else (
    echo Deploy failed -- is SolidWorks still running?
)
