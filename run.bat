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

REM 源码模式用系统 Python(可能 3.14 + pywebview 6.x)走浏览器；打包 exe 用内置 pywebview 5.4 独立窗口
set TABEXPORTER_NO_WEBVIEW=1

%PY% launcher.py %*
pause
