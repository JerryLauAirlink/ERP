from datetime import datetime

from sqlalchemy import BigInteger, Boolean, Column, DateTime, ForeignKey, Integer, Numeric, String, Text
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy import JSON
from sqlalchemy.orm import relationship

from database import Base

# SQLite uses JSON; PostgreSQL uses JSONB when available
JsonType = JSON().with_variant(JSONB, "postgresql")


class Tenant(Base):
    __tablename__ = "tenants"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, index=True)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime)

    users = relationship("User", back_populates="tenant")


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    tenant_id = Column(Integer, ForeignKey("tenants.id"))
    email = Column(String, index=True)
    hashed_password = Column(String)
    name = Column(String)
    role = Column(String)
    is_active = Column(Boolean, default=True)
    permissions = Column(JsonType)
    allowed_regions = Column(JsonType)

    tenant = relationship("Tenant", back_populates="users")


class ErpBackup(Base):
    """Full ERP snapshot (same structure as Settings → Download Backup JSON)."""

    __tablename__ = "erp_backups"

    id = Column(Integer, primary_key=True, index=True)
    tenant_id = Column(Integer, ForeignKey("tenants.id"), index=True)
    version = Column(Integer, default=1)
    exported_at = Column(DateTime, default=datetime.utcnow)
    payload = Column(JsonType, nullable=False)
    note = Column(String)


class Customer(Base):
    __tablename__ = "customers"

    id = Column(Integer, primary_key=True, index=True)
    tenant_id = Column(Integer, ForeignKey("tenants.id"))
    name = Column(String, index=True)
    type = Column(String, default="CUSTOMER")

    tenant = relationship("Tenant")


class ARInvoice(Base):
    __tablename__ = "ar_invoices"

    id = Column(Integer, primary_key=True, index=True)
    tenant_id = Column(Integer, ForeignKey("tenants.id"))
    customer_id = Column(Integer, ForeignKey("customers.id"), nullable=True)
    cust_po = Column(String)
    po_completed = Column(Boolean, default=False)
    invoice_number = Column(String)
    currency = Column(String)
    invoice_amount = Column(Numeric)
    amount_hkd = Column(Numeric)
    invoice_date = Column(DateTime)
    due_date = Column(DateTime)
    payment_received_date = Column(DateTime, nullable=True)
    po_type = Column(String)
    remark = Column(Text)
    balance = Column(Numeric)
    status = Column(String)

    tenant = relationship("Tenant")
    customer = relationship("Customer")


class ExchangeRate(Base):
    __tablename__ = "exchange_rates"

    id = Column(Integer, primary_key=True, index=True)
    tenant_id = Column(Integer, ForeignKey("tenants.id"))
    from_currency = Column(String)
    to_currency = Column(String, default="HKD")
    rate = Column(Numeric)
    updated_at = Column(DateTime)

    tenant = relationship("Tenant")


class ErpEntityRecord(Base):
    __tablename__ = "erp_entity_records"

    id = Column(BigInteger, primary_key=True, index=True)
    tenant_id = Column(Integer, ForeignKey("tenants.id"), index=True)
    entity_type = Column(String(64), index=True)
    entity_id = Column(BigInteger, index=True)
    region = Column(String(16))
    payload = Column(JsonType, nullable=False, default=dict)
    is_deleted = Column(Boolean, default=False)
    updated_at = Column(DateTime, default=datetime.utcnow)
    updated_by = Column(BigInteger)


class ErpSyncMeta(Base):
    __tablename__ = "erp_sync_meta"

    tenant_id = Column(Integer, ForeignKey("tenants.id"), primary_key=True)
    server_version = Column(BigInteger, default=0)
    updated_at = Column(DateTime, default=datetime.utcnow)
