@echo off
REM Rebuild AIRLINK_GITHUB upload folder from project source
set SRC=%~dp0..
set DST=%~dp0

echo Rebuilding AIRLINK_GITHUB pack...

if exist "%DST%AIRLINK" rmdir /s /q "%DST%AIRLINK"
if exist "%DST%api" rmdir /s /q "%DST%api"
if exist "%DST%supabase" rmdir /s /q "%DST%supabase"
if exist "%DST%scripts" rmdir /s /q "%DST%scripts"

mkdir "%DST%AIRLINK"
mkdir "%DST%api"
mkdir "%DST%supabase\migrations"
mkdir "%DST%scripts"

copy /Y "%SRC%\AIRLINK\index.html" "%DST%AIRLINK\index.html"
copy /Y "%SRC%\AIRLINK\logo.png" "%DST%AIRLINK\"
copy /Y "%SRC%\AIRLINK\login-logo.png" "%DST%AIRLINK\"
copy /Y "%SRC%\AIRLINK\login-screen.png" "%DST%AIRLINK\"
copy /Y "%SRC%\api\index.py" "%DST%api\"
copy /Y "%SRC%\api\requirements.txt" "%DST%api\"
copy /Y "%SRC%\database.py" "%DST%\"
copy /Y "%SRC%\models.py" "%DST%\"
copy /Y "%SRC%\erp_backup.py" "%DST%\"
copy /Y "%SRC%\erp_live_sync.py" "%DST%\"
copy /Y "%SRC%\requirements.txt" "%DST%\"
copy /Y "%SRC%\runtime.txt" "%DST%\"
copy /Y "%SRC%\world-map-data.js" "%DST%\"
copy /Y "%SRC%\supabase\migrations\001_initial_schema.sql" "%DST%supabase\migrations\"
copy /Y "%SRC%\supabase\migrations\002_live_sync_entities.sql" "%DST%supabase\migrations\"
copy /Y "%SRC%\scripts\SETUP_SUPABASE.bat" "%DST%scripts\"
copy /Y "%SRC%\scripts\seed_supabase.py" "%DST%scripts\"

echo Done. Upload everything INSIDE AIRLINK_GITHUB to GitHub repo root.
pause
