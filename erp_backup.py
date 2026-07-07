import os
from datetime import datetime

from fastapi import Header, HTTPException, status
from sqlalchemy.orm import Session

from models import ErpBackup, Tenant


def default_tenant_id() -> int:
    return int(os.getenv("ERP_TENANT_ID", "1"))


def verify_sync_secret(x_erp_sync_key: str | None = Header(default=None, alias="X-ERP-Sync-Key")) -> None:
    expected = os.getenv("ERP_SYNC_SECRET", "").strip()
    if not expected:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Cloud sync not configured (ERP_SYNC_SECRET missing on server)",
        )
    if not x_erp_sync_key or x_erp_sync_key != expected:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Invalid sync key")


def ensure_tenant(db: Session, tenant_id: int) -> Tenant:
    tenant = db.query(Tenant).filter(Tenant.id == tenant_id).first()
    if tenant:
        return tenant
    tenant = Tenant(id=tenant_id, name="Default Company", is_active=True, created_at=datetime.utcnow())
    db.add(tenant)
    db.commit()
    db.refresh(tenant)
    return tenant


def save_erp_backup(db: Session, tenant_id: int, payload: dict, note: str | None = None) -> ErpBackup:
    ensure_tenant(db, tenant_id)
    version = int(payload.get("version") or 1)
    exported_raw = payload.get("exportedAt")
    exported_at = datetime.fromisoformat(exported_raw.replace("Z", "+00:00")) if exported_raw else datetime.utcnow()
    record = ErpBackup(
        tenant_id=tenant_id,
        version=version,
        exported_at=exported_at,
        payload=payload,
        note=note,
    )
    db.add(record)
    db.commit()
    db.refresh(record)
    return record


def load_latest_erp_backup(db: Session, tenant_id: int) -> ErpBackup | None:
    return (
        db.query(ErpBackup)
        .filter(ErpBackup.tenant_id == tenant_id)
        .order_by(ErpBackup.exported_at.desc(), ErpBackup.id.desc())
        .first()
    )
