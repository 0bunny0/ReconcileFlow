@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install Python and enable Add Python to PATH.
  exit /b 1
)

python -c "import pandas, openpyxl, tzdata" >nul 2>nul
if not errorlevel 1 (
  echo Python dependencies are already installed.
  exit /b 0
)

python -c "import platform,sys; raise SystemExit(0 if sys.version_info[:2] == (3,14) and platform.architecture()[0] == '64bit' else 1)" >nul 2>nul
if errorlevel 1 goto ONLINE_INSTALL

if not exist "%~dp0wheels\pandas-3.0.5-cp314-cp314-win_amd64.whl" goto ONLINE_INSTALL

echo Installing Python dependencies from the offline package...
python -m pip install --disable-pip-version-check --no-index --find-links="%~dp0wheels" -r "%~dp0requirements-offline-cp314-win64.txt"
if errorlevel 1 (
  echo Offline dependency installation failed.
  exit /b 1
)

python -c "import pandas, openpyxl, tzdata" >nul 2>nul
if errorlevel 1 (
  echo Dependencies were installed but could not be imported.
  exit /b 1
)

echo Offline dependency installation completed.
exit /b 0

:ONLINE_INSTALL
echo No matching offline wheels were found for this Python version.
echo Trying the normal online installation...
python -m pip install --disable-pip-version-check -r "%~dp0requirements.txt"
if errorlevel 1 (
  echo Online installation failed. This offline package targets Python 3.14 64-bit.
  exit /b 1
)
exit /b 0
