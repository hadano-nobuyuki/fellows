@echo off
rem Training environment setup
net session >nul 2>&1
if %errorlevel% neq 0 (
  if "%~1"=="elevated" (
    echo.
    echo 管理者権限で実行できませんでした。処理を中止します。
    echo 確認画面で［いいえ］を押したか、管理者のパスワードが違っていたようです。
    echo setup.bat を右クリックして［管理者として実行］で試してください。
    echo.
    pause
    exit /b 1
  )
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -ArgumentList 'elevated' -Verb RunAs"
  exit /b
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1"
pause
