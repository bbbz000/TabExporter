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

echo [1/3] Installing dependencies...
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Dependency install failed.
    pause
    exit /b 1
)

echo [2/3] Building single-file exe...
%PY% -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name ExportTables ^
  --add-data "gui;gui" ^
  --add-data "tabtoy.exe;." ^
  --collect-all webview ^
  app.py
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo [3/3] Done! Output: dist\ExportTables.exe
pause
