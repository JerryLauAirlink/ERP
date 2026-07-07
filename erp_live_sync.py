"""Per-entity live sync for multi-user ERP (Option D)."""
from datetime import datetime

from sqlalchemy.orm import Session

from models import ErpEntityRecord, ErpSyncMeta


VALID_ENTITY_TYPES = frozenset({
    "clients",
    "jobs",
    "quotations",
    "vendors",
    "sis",
    "ar_invoices",
    "ap_bills",
    "users",
    "audit_logs",
    "monthly_po_lines",
    "monthly_ar_lines",
    "monthly_ar_expected",
    "settings",
})


def _ensure_sync_meta(db: Session, tenant_id: int) -> ErpSyncMeta:
    meta = db.query(ErpSyncMeta).filter(ErpSyncMeta.tenant_id == tenant_id).first()
    if meta:
        return meta
    meta = ErpSyncMeta(tenant_id=tenant_id, server_version=0, updated_at=datetime.utcnow())
    db.add(meta)
    db.commit()
    db.refresh(meta)
    return meta


def get_sync_status(db: Session, tenant_id: int) -> dict:
    meta = _ensure_sync_meta(db, tenant_id)
    count = db.query(ErpEntityRecord).filter(
        ErpEntityRecord.tenant_id == tenant_id,
        ErpEntityRecord.is_deleted.is_(False),
    ).count()
    return {
        "server_version": int(meta.server_version or 0),
        "updated_at": meta.updated_at.isoformat() if meta.updated_at else None,
        "entity_count": count,
    }


def bump_version(db: Session, tenant_id: int) -> int:
    meta = _ensure_sync_meta(db, tenant_id)
    meta.server_version = int(meta.server_version or 0) + 1
    meta.updated_at = datetime.utcnow()
    db.commit()
    db.refresh(meta)
    return int(meta.server_version)


def upsert_entity_changes(
    db: Session,
    tenant_id: int,
    changes: list[dict],
    updated_by: int | None = None,
) -> int:
    if not changes:
        return get_sync_status(db, tenant_id)["server_version"]

    now = datetime.utcnow()
    for ch in changes:
        entity_type = str(ch.get("entity_type") or "").strip()
        if entity_type not in VALID_ENTITY_TYPES:
            continue
        entity_id = int(ch.get("entity_id") or 0)
        is_deleted = bool(ch.get("is_deleted"))
        payload = ch.get("payload") if isinstance(ch.get("payload"), dict) else {}
        region = ch.get("region") or payload.get("region")

        row = (
            db.query(ErpEntityRecord)
            .filter(
                ErpEntityRecord.tenant_id == tenant_id,
                ErpEntityRecord.entity_type == entity_type,
                ErpEntityRecord.entity_id == entity_id,
            )
            .first()
        )
        if row:
            row.payload = payload
            row.is_deleted = is_deleted
            row.region = region
            row.updated_at = now
            row.updated_by = updated_by
        else:
            db.add(
                ErpEntityRecord(
                    tenant_id=tenant_id,
                    entity_type=entity_type,
                    entity_id=entity_id,
                    region=region,
                    payload=payload,
                    is_deleted=is_deleted,
                    updated_at=now,
                    updated_by=updated_by,
                )
            )

    db.commit()
    return bump_version(db, tenant_id)


def get_changes_since(db: Session, tenant_id: int, since_version: int) -> dict:
    meta = _ensure_sync_meta(db, tenant_id)
    current = int(meta.server_version or 0)
    if since_version >= current:
        return {"server_version": current, "changes": []}

    rows = (
        db.query(ErpEntityRecord)
        .filter(ErpEntityRecord.tenant_id == tenant_id)
        .order_by(ErpEntityRecord.updated_at.asc(), ErpEntityRecord.id.asc())
        .all()
    )
    changes = [
        {
            "entity_type": r.entity_type,
            "entity_id": r.entity_id,
            "region": r.region,
            "payload": r.payload or {},
            "is_deleted": bool(r.is_deleted),
            "updated_at": r.updated_at.isoformat() if r.updated_at else None,
            "updated_by": r.updated_by,
        }
        for r in rows
    ]
    return {"server_version": current, "changes": changes}


def load_full_entities(db: Session, tenant_id: int) -> dict:
    rows = (
        db.query(ErpEntityRecord)
        .filter(
            ErpEntityRecord.tenant_id == tenant_id,
            ErpEntityRecord.is_deleted.is_(False),
        )
        .all()
    )
    grouped: dict[str, list] = {t: [] for t in VALID_ENTITY_TYPES}
    singletons: dict[str, dict] = {}
    for r in rows:
        if r.entity_type in ("monthly_ar_expected", "settings"):
            singletons[r.entity_type] = r.payload or {}
        elif r.entity_type in grouped:
            grouped[r.entity_type].append(r.payload or {})

    meta = _ensure_sync_meta(db, tenant_id)
    return {
        "server_version": int(meta.server_version or 0),
        "entities": grouped,
        "singletons": singletons,
    }
