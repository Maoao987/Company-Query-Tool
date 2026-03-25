@echo off
setlocal
cd /d "%~dp0"

if /I not "%~1"=="--foreground" (
    if exist "%~dp0start_hidden.vbs" (
        wscript.exe "%~dp0start_hidden.vbs"
        exit /b
    )
)

if exist ".venv\Scripts\streamlit.exe" (
    start "" http://localhost:8501
    ".venv\Scripts\streamlit.exe" run "%~dp0app.py" --server.headless true --browser.gatherUsageStats false
    if /I "%~1"=="--foreground" exit /b %ERRORLEVEL%
    pause
) else (
    echo.
    echo  [!] Please run "Install.bat" first!
    echo      Please double-click "Install.bat" in this folder.
    echo.
    if /I "%~1"=="--foreground" exit /b 1
    pause
)
