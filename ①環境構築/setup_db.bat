@echo off
rem ------------------------------------------------------------
rem  Create database only (for PCs that already have PostgreSQL)
rem ------------------------------------------------------------
net session >nul 2>&1
if %errorlevel% equ 0 goto run
rem 管理者として起動し直す(1回だけ。起動し直しても管理者になれなければ停止する)
if "%~1"=="elevated" goto noadmin
powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -ArgumentList 'elevated' -Verb RunAs"
exit /b
:noadmin
echo.
echo 管理者権限で実行できませんでした。処理を中止します。
echo ・確認画面で［いいえ］を押した、または管理者のパスワードが違う
echo ・この PC では管理者への切り替えができない
echo のどちらかです。setup.bat を右クリックして［管理者として実行］で試してください。
echo.
pause
exit /b 1
:run
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" -DbOnly
pause
