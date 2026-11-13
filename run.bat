@echo off
cd /d "%~dp0"

set "PY=python"
where python >nul 2>nul
if errorlevel 1 (
    set "PY=py"
    where py >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] Python 3.10+ not found. Please install it and check "Add to PATH".
        pause
        exit /b 1
    )
)

%PY% launcher.py %*
pause
