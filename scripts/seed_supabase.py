#!/usr/bin/env python3
"""Seed demo ERP backup into Supabase. Usage:
  1. Copy .env.example → .env and set DATABASE_URL + ERP_SYNC_SECRET
  2. Run SQL in supabase/migrations/001_initial_schema.sql first
  3. python scripts/seed_supabase.py
"""
import json
import os
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

try:
    from dotenv import load_dotenv
    load_dotenv(ROOT / ".env")
except ImportError:
    pass

from database import SessionLocal, check_db_connection
from erp_backup import default_tenant_id, save_erp_backup

DEMO_PAYLOAD = {
    "version": 1,
    "exportedAt": datetime.utcnow().isoformat() + "Z",
    "app": "ERP System",
    "data": {
        "clients": [
            {"id": 1, "region": "HK", "customer_no": "CUST-0001", "company": "Pacific Systems", "gst_no": "GST-HK-1102", "primary_contact": "Ken Wong", "company_phone": "852-2333-1122", "mobile_phone": "852-9456-3321", "email": "ar@pacific.com", "address": "18/F Pacific Tower, TST, Hong Kong", "postal_code": "999077", "account_dept_contact": "May Chan", "payment_terms": "30 Days"},
            {"id": 2, "region": "HK", "customer_no": "CUST-0002", "company": "Kingston Trading", "gst_no": "GST-HK-2208", "primary_contact": "Henry Lee", "company_phone": "852-2550-0091", "mobile_phone": "852-9877-1234", "email": "finance@kingston.com", "address": "Rm 1206, Mongkok Centre, Kowloon", "postal_code": "999078", "account_dept_contact": "Amy Lau", "payment_terms": "45 Days"},
        ],
        "jobs": [],
        "quotations": [],
        "vendors": [],
        "sis": [],
        "arInvoices": [],
        "apBills": [],
        "users": [],
        "auditLogs": [],
        "monthlyPoLines": [],
        "monthlyArLines": [],
        "monthlyArExpectedSnapshots": {},
        "settings": {"lang": "zh_TW", "activeRegion": "HK"},
    },
}


def main():
    status = check_db_connection()
    if not status.get("ok"):
        print("Database connection failed:", status.get("error"))
        sys.exit(1)
    print("Connected:", status.get("provider"))
    db = SessionLocal()
    try:
        record = save_erp_backup(db, default_tenant_id(), DEMO_PAYLOAD, note="seed script")
        print("Saved backup id=", record.id)
    finally:
        db.close()


if __name__ == "__main__":
    main()
