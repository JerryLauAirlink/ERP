@echo off
chcp 65001 >nul
echo.
echo ============================================
echo   ERP + Supabase 連線設定（2026 新版介面）
echo ============================================
echo.
echo [已完成] SQL 建表（Success. No rows returned）
echo.
echo --- 攞 DATABASE_URL（兩個方法任揀一）---
echo.
echo 方法 A（最簡單）:
echo   1. 打開 Supabase 專案主頁（唔係 Settings 齒輪入面）
echo   2. 撳頂部綠色「Connect」掣
echo   3. 揀 URI → Transaction pooler（port 6543）
echo   4. Copy，把 [YOUR-PASSWORD] 換成真密碼
echo.
echo 方法 B:
echo   1. 左邊主 sidebar 撳「Database」（圖示係圓柱體，唔係 Settings）
echo   2. 入面搵 Configuration 或 Connect
echo   3. 複製 Transaction pooler URI
echo.
echo 你嘅 project 範本（只換密碼）:
echo postgresql://postgres.qnmhupzvtadydiastzu:密碼@aws-0-ap-southeast-2.pooler.supabase.com:6543/postgres
echo.
echo --- Vercel 環境變數 ---
echo   DATABASE_URL  = 上面完整 URI
echo   ERP_SYNC_SECRET = 自己設（例如 MyErp2026Secret）
echo   SECRET_KEY    = 隨機字串
echo   然後 Redeploy
echo.
echo --- ERP 上傳資料 ---
echo   設定 - 備份 - 輸入同步金鑰 - 上傳至雲端
echo.
echo --- 即時多人同步（Option D）---
echo   1. Supabase SQL Editor 執行 supabase/migrations/002_live_sync_entities.sql
echo   2. git push 部署最新 code
echo   3. 設定 - 備份 - 勾選「啟用即時同步」
echo.
echo 驗證: https://你的域名/api/health
echo.
pause
