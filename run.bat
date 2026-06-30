@echo off
cd /d "%~dp0"
echo 正在啟動 ERP API...
echo.
echo  UI 頁面:  http://127.0.0.1:8080/ui
echo  API 文件: http://127.0.0.1:8080/docs
echo.
uvicorn main:app --host 127.0.0.1 --port 8080
pause
