AIRLINK GitHub Upload Pack
==========================

HOW TO UPLOAD (important!)
--------------------------
1. Open GitHub repo: JerryLauAirlink/ERP
2. Delete ALL old files/folders in the repo first (start clean)
   OR upload these files to REPLACE matching paths
3. Open folder: AIRLINK_GITHUB on your PC
4. Select ALL files and folders INSIDE AIRLINK_GITHUB
   (NOT the AIRLINK_GITHUB folder itself)
5. Drag to GitHub → "Add file" → "Upload files"
6. Commit message: AIRLINK deploy pack
7. Wait Vercel deploy (2-3 min)
8. Open: https://airlink-erp.vercel.app
9. Press F12 → Console → must see:
   AIRLINK ERP build: airlink-2026-07-09e

SUPABASE SQL (required once — fixes import sync 500 error)
----------------------------------------------------------
Open Supabase Dashboard → SQL Editor → run:
  supabase/migrations/003_bigint_entity_ids.sql

This fixes "integer out of range" when syncing imported records.

FOLDER STRUCTURE (upload to GitHub ROOT)
----------------------------------------
AIRLINK_GITHUB/
  index.html              ← root redirect (fixes 404)
  vercel.json
  runtime.txt
  requirements.txt
  world-map-data.js
  database.py
  models.py
  erp_backup.py
  erp_live_sync.py
  AIRLINK/
    index.html            ← main shell (loads erp-app.jsx)
    erp-app.jsx           ← app code (required!)
    logo.png
    login-logo.png
    login-screen.png
  api/
    index.py
    requirements.txt
  supabase/migrations/
    001_initial_schema.sql
    002_live_sync_entities.sql
    003_bigint_entity_ids.sql
  scripts/
    SETUP_SUPABASE.bat
    seed_supabase.py

DO NOT UPLOAD
-------------
- DEMO/  DEMO_INTL/  node_modules/
- auth.py  main.py  services.py  package.json
- public/ folder (not needed)

VERCEL ENV (must be set in Vercel dashboard)
--------------------------------------------
DATABASE_URL = your Supabase postgres URL
ERP_SYNC_SECRET = your sync secret

After login as root:
Settings → Backup → enter sync key → live sync on
