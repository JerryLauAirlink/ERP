from sqlalchemy import Boolean, Column, DateTime, ForeignKey, Integer, Numeric, String
from sqlalchemy.orm import relationship

from database import Base


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
    email = Column(String, unique=True, index=True)
    hashed_password = Column(String)
    role = Column(String)
    is_active = Column(Boolean, default=True)

    tenant = relationship("Tenant", back_populates="users")


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
    remark = Column(String)
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
