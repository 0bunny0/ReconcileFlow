@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

call "%~dp0install-dependencies.bat"
if errorlevel 1 (
  echo.
  echo Startup stopped because dependencies could not be installed.
  pause
  exit /b 1
)

python app.py --host 0.0.0.0 --port 8765 --no-browser
pause
