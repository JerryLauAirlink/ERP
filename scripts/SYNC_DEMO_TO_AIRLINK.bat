@echo off
chcp 65001 >nul
echo.
echo ============================================
echo   DEMO -^> AIRLINK 更新（俾客戶用）
echo ============================================
echo.
echo 用途：你在 DEMO 改完功能後，用此腳本更新 AIRLINK 客戶版。
echo.
echo 步驟：
echo   1. 平時只在 DEMO\index.html 加改功能（/demo 開發版，免 Login）
echo   2. 測試 OK 後執行此腳本
echo   3. git push 部署 — 客戶開主網址見 AIRLINK 版
echo.
set /p CONFIRM=確定要將 DEMO 複製到 AIRLINK？（AIRLINK 會保留 Login+品牌）[Y/N]:
if /I not "%CONFIRM%"=="Y" exit /b 0

copy /Y "DEMO\index.html" "AIRLINK\index.html"
if errorlevel 1 (
  echo 複製失敗
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0apply-airlink-client.ps1"
if errorlevel 1 (
  echo AIRLINK 客戶設定套用失敗
  exit /b 1
)

copy /Y "AIRLINK\index.html" "index.html"
echo.
echo 完成！
echo   開發版： https://你的域名/demo
echo   客戶版： https://你的域名/
echo.
pause
