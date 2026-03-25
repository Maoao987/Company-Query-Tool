@echo off
echo.
echo  Company Query Tool - Installer
echo  ================================
echo.

if not exist "%~dp0install.ps1" (
    echo  [ERROR] Missing install.ps1 in this folder.
    echo  Please EXTRACT the ZIP first, then run Install.bat.
    pause
    exit /b 1
)

:: Copy ps1 to TEMP (ASCII path) to work around Chinese folder names
copy /Y "%~dp0install.ps1" "%TEMP%\cqt_setup.ps1" >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Cannot copy installer.
    echo  Please EXTRACT the ZIP first, then run Install.bat.
    pause
    exit /b 1
)

:: Run PowerShell from TEMP; pass real app dir as -AppDir parameter
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%TEMP%\cqt_setup.ps1" -AppDir "%~dp0"
set _ERR=%ERRORLEVEL%
del "%TEMP%\cqt_setup.ps1" >nul 2>&1

if %_ERR% NEQ 0 (
    echo.
    echo  [ERROR] Installer failed. Code: %_ERR%
    echo  Please screenshot this window and contact support.
    pause
)
