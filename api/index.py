"""Vercel serverless entry — lightweight API (no auth/passlib import chain)."""
import json
import os
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from fastapi import Depends, FastAPI, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from sqlalchemy.orm import Session

from database import check_db_connection, get_db
from erp_backup import (
    default_tenant_id,
    load_latest_erp_backup,
    save_erp_backup,
    verify_sync_secret,
)
from erp_live_sync import (
    get_changes_since,
    get_sync_status,
    load_full_entities,
    upsert_entity_changes,
)

app = FastAPI(title="ERP Cloud API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class ErpBackupPayload(BaseModel):
    version: int = 1
    exportedAt: str | None = None
    app: str | None = None
    data: dict


class EntityChange(BaseModel):
    entity_type: str
    entity_id: int
    payload: dict = {}
    is_deleted: bool = False
    region: str | None = None


class EntityChangesPayload(BaseModel):
    changes: list[EntityChange]
    updated_by: int | None = None


@app.get("/")
@app.get("/api")
def api_root():
    return {"status": "ok", "message": "ERP API"}


# Daily FX cache (module-level; resets on cold start — client also caches by UTC date)
_FX_CACHE: dict = {"date": None, "payload": None}
_FX_CURRENCIES = ("USD", "HKD", "SGD", "TWD", "RMB", "EUR", "JPY", "MYR", "IDR", "AUD", "THB")
_FX_FALLBACK = {
    "USD": 1,
    "HKD": 7.8,
    "SGD": 1.35,
    "TWD": 32.2,
    "RMB": 7.2,
    "EUR": 0.92,
    "JPY": 160,
    "MYR": 4.7,
    "IDR": 16000,
    "AUD": 1.55,
    "THB": 36.5,
}


def _utc_date() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d")


def _normalize_usd_rates(raw_rates: dict) -> dict:
    """Map provider rates (USD base) into AIRLINK currency codes. CNY → RMB."""
    merged = dict(raw_rates or {})
    if "RMB" not in merged and merged.get("CNY") is not None:
        merged["RMB"] = merged["CNY"]
    out = {}
    for code in _FX_CURRENCIES:
        try:
            val = float(merged.get(code, _FX_FALLBACK[code]))
        except (TypeError, ValueError):
            val = float(_FX_FALLBACK[code])
        if val <= 0:
            val = float(_FX_FALLBACK[code])
        out[code] = val
    out["USD"] = 1.0
    return out


def _fetch_open_er_api() -> dict:
    url = "https://open.er-api.com/v6/latest/USD"
    req = urllib.request.Request(url, headers={"User-Agent": "AIRLINK-ERP/1.0", "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=12) as resp:
        body = json.loads(resp.read().decode("utf-8"))
    if body.get("result") != "success" or not isinstance(body.get("rates"), dict):
        raise RuntimeError("FX provider returned an unexpected payload")
    rates = _normalize_usd_rates(body["rates"])
    return {
        "ok": True,
        "base": "USD",
        "date": _utc_date(),
        "updated_at": body.get("time_last_update_utc") or datetime.now(timezone.utc).isoformat(),
        "next_update_at": body.get("time_next_update_utc"),
        "source": "exchangerate-api.com",
        "provider": body.get("provider") or "https://www.exchangerate-api.com",
        "rates": rates,
    }


@app.get("/fx")
@app.get("/api/fx")
def get_fx_rates(force: int = 0):
    """Daily market FX vs USD. Cached per UTC day. Attribution: exchangerate-api.com."""
    today = _utc_date()
    if not force and _FX_CACHE.get("date") == today and _FX_CACHE.get("payload"):
        return _FX_CACHE["payload"]
    try:
        payload = _fetch_open_er_api()
        _FX_CACHE["date"] = today
        _FX_CACHE["payload"] = payload
        return payload
    except Exception as exc:
        cached = _FX_CACHE.get("payload")
        if cached:
            return {**cached, "ok": True, "stale": True, "error": str(exc)}
        return {
            "ok": False,
            "base": "USD",
            "date": today,
            "updated_at": None,
            "source": "fallback",
            "rates": dict(_FX_FALLBACK),
            "error": str(exc),
        }


@app.get("/health")
@app.get("/api/health")
def health_check():
    db_status = check_db_connection()
    return {
        "status": "ok" if db_status.get("ok") else "degraded",
        "api": "ERP Cloud API",
        "database": db_status,
        "cloud_sync_configured": bool(os.getenv("ERP_SYNC_SECRET")),
        "live_sync_ready": True,
    }


@app.get("/backup")
@app.get("/api/backup")
def get_cloud_backup(
    db: Session = Depends(get_db),
    _: None = Depends(verify_sync_secret),
    tenant_id: int = default_tenant_id(),
):
    record = load_latest_erp_backup(db, tenant_id)
    if not record:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="No cloud backup found")
    return {
        "id": record.id,
        "tenant_id": record.tenant_id,
        "exported_at": record.exported_at.isoformat() if record.exported_at else None,
        "payload": record.payload,
    }


@app.post("/backup")
@app.post("/api/backup")
def post_cloud_backup(
    body: ErpBackupPayload,
    db: Session = Depends(get_db),
    _: None = Depends(verify_sync_secret),
    tenant_id: int = default_tenant_id(),
):
    if not body.data or not isinstance(body.data, dict):
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Invalid backup: missing data")
    if not isinstance(body.data.get("clients"), list) or not isinstance(body.data.get("jobs"), list):
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Invalid backup: clients/jobs required")
    record = save_erp_backup(db, tenant_id, body.model_dump(), note="cloud sync")
    users = body.data.get("users") or []
    if isinstance(users, list) and users:
        changes = []
        for u in users:
            if not isinstance(u, dict) or u.get("id") is None:
                continue
            changes.append({
                "entity_type": "users",
                "entity_id": int(u["id"]),
                "payload": u,
                "is_deleted": False,
                "region": None,
            })
        if changes:
            upsert_entity_changes(db, tenant_id, changes, updated_by=None)
    return {
        "message": "Cloud backup saved",
        "id": record.id,
        "exported_at": record.exported_at.isoformat() if record.exported_at else None,
    }


@app.get("/sync/client-config")
@app.get("/api/sync/client-config")
def sync_client_config():
    """Auto-sync config for ERP UI (internal team — key delivered to browser)."""
    secret = os.getenv("ERP_SYNC_SECRET", "").strip()
    if not secret:
        return {"auto_sync": False, "sync_key": None, "pull_public": False}
    return {
        "auto_sync": True,
        "sync_key": secret,
        "pull_public": True,
    }


@app.get("/sync/status")
@app.get("/api/sync/status")
def sync_status(
    db: Session = Depends(get_db),
    tenant_id: int = default_tenant_id(),
):
    try:
        return get_sync_status(db, tenant_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/sync/changes")
@app.get("/api/sync/changes")
def sync_changes(
    since: int = 0,
    db: Session = Depends(get_db),
    tenant_id: int = default_tenant_id(),
):
    try:
        return get_changes_since(db, tenant_id, since)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.get("/sync/full")
@app.get("/api/sync/full")
def sync_full(
    db: Session = Depends(get_db),
    tenant_id: int = default_tenant_id(),
):
    try:
        return load_full_entities(db, tenant_id)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/sync/push")
@app.post("/api/sync/push")
def sync_push(
    body: EntityChangesPayload,
    db: Session = Depends(get_db),
    _: None = Depends(verify_sync_secret),
    tenant_id: int = default_tenant_id(),
):
    if not body.changes:
        return {"server_version": get_sync_status(db, tenant_id)["server_version"], "applied": 0}
    try:
        changes = [c.model_dump() for c in body.changes]
        version = upsert_entity_changes(db, tenant_id, changes, updated_by=body.updated_by)
        return {"server_version": version, "applied": len(changes)}
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
