"""Vercel serverless entry — lightweight API (no auth/passlib import chain)."""
import os
import sys

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
    return {
        "message": "Cloud backup saved",
        "id": record.id,
        "exported_at": record.exported_at.isoformat() if record.exported_at else None,
    }


@app.get("/sync/status")
@app.get("/api/sync/status")
def sync_status(
    db: Session = Depends(get_db),
    _: None = Depends(verify_sync_secret),
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
    _: None = Depends(verify_sync_secret),
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
    _: None = Depends(verify_sync_secret),
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
