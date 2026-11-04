@echo off
rem ------------------------------------------------------------
rem  Create database only (for PCs that already have PostgreSQL)
rem ------------------------------------------------------------
net session >nul 2>&1
if %errorlevel% neq 0 (
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" -DbOnly
pause
