@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_CMD="
python --version >nul 2>nul && set "PYTHON_CMD=python"
if not defined PYTHON_CMD (
  py -3 --version >nul 2>nul && set "PYTHON_CMD=py -3"
)

if not defined PYTHON_CMD (
  echo [ERROR] Python 3 was not found in PATH.
  pause
  exit /b 1
)

set "LOG_FILE=%~dp0startup.log"
echo ==== %date% %time% ==== > "%LOG_FILE%"

echo [1/3] Installing requirements...
call %PYTHON_CMD% -m pip install -r requirements.txt >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
  echo [ERROR] Failed to install requirements.
  echo Check startup.log for details.
  pause
  exit /b 1
)

echo [2/3] Starting server...
start "Bid Parser Demo" /D "%~dp0" cmd /k "%PYTHON_CMD% app.py"
if errorlevel 1 (
  echo [ERROR] Failed to open server window.
  pause
  exit /b 1
)

echo [3/3] Opening browser...
timeout /t 3 /nobreak >nul
start "" http://127.0.0.1:8008

echo Done. If something fails later, check startup.log.
exit /b 0
