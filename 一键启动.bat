@echo off
setlocal
cd /d "%~dp0"
set "START_ARGS="
if defined DRY_RUN set "START_ARGS=-DryRun"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0start.ps1" %START_ARGS%
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
  echo.
  echo Startup failed. See startup.log for details.
  pause
)
exit /b %EXIT_CODE%
