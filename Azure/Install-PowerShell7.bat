@echo off
echo ================================================
echo   PowerShell 7+ Installation Check
echo ================================================
echo.

REM Check if pwsh.exe exists
where pwsh.exe >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [OK] PowerShell 7+ is already installed!
    pwsh.exe -Command "Write-Host 'Version: ' -NoNewline; $PSVersionTable.PSVersion.ToString()"
    echo.
    echo You can now run Start-ServicedeskTools.bat
    pause
    exit /b 0
)

echo [!] PowerShell 7+ is not installed.
echo.
echo This script will attempt to install PowerShell 7+ using one of the following methods:
echo   1. winget (Windows Package Manager)
echo   2. Direct download from Microsoft
echo.

REM Check if winget is available
where winget.exe >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [INFO] Found winget. Installing PowerShell 7 via winget...
    echo.
    winget install --id Microsoft.PowerShell --source winget --silent --accept-package-agreements --accept-source-agreements
    
    if %ERRORLEVEL% EQU 0 (
        echo.
        echo [SUCCESS] PowerShell 7+ has been installed successfully!
        echo Please close this window and run Start-ServicedeskTools.bat
        pause
        exit /b 0
    ) else (
        echo.
        echo [ERROR] Installation via winget failed.
        goto MANUAL_INSTALL
    )
) else (
    echo [!] winget is not available.
    goto MANUAL_INSTALL
)

:MANUAL_INSTALL
echo.
echo ================================================
echo   Manual Installation Required
echo ================================================
echo.
echo Please install PowerShell 7+ manually:
echo.
echo 1. Visit: https://github.com/PowerShell/PowerShell/releases/latest
echo 2. Download: PowerShell-7.x.x-win-x64.msi
echo 3. Run the installer
echo 4. After installation, run Start-ServicedeskTools.bat
echo.
echo Alternatively, you can run this command in PowerShell 5.1:
echo.
echo   iex "& { $(irm https://aka.ms/install-powershell.ps1) } -UseMSI"
echo.
pause
exit /b 1
