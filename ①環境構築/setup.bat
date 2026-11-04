@echo off
rem ------------------------------------------------------------
rem  Training environment setup (runs setup.ps1 as administrator)
rem ------------------------------------------------------------
net session >nul 2>&1
if %errorlevel% neq 0 (
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1"
pause
