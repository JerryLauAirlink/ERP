-- ERP System — Supabase / PostgreSQL initial schema
-- Run in Supabase Dashboard → SQL Editor → New query → Run

-- Extensions
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- Tenants (multi-company ready; default tenant id = 1)
CREATE TABLE IF NOT EXISTS tenants (
  id SERIAL PRIMARY KEY,
  name VARCHAR(255) NOT NULL,
  is_active BOOLEAN NOT NULL DEFAULT TRUE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

INSERT INTO tenants (id, name, is_active, created_at)
VALUES (1, 'Default Company', TRUE, NOW())
ON CONFLICT (id) DO NOTHING;

-- Users (API / future auth)
CREATE TABLE IF NOT EXISTS users (
  id SERIAL PRIMARY KEY,
  tenant_id INTEGER NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  email VARCHAR(255) NOT NULL,
  hashed_password VARCHAR(255),
  name VARCHAR(255),
  role VARCHAR(64) NOT NULL DEFAULT 'user',
  is_active BOOLEAN NOT NULL DEFAULT TRUE,
  permissions JSONB,
  allowed_regions JSONB,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (tenant_id, email)
);

-- Full ERP snapshot (mirrors browser backup JSON — easiest migration path)
CREATE TABLE IF NOT EXISTS erp_backups (
  id SERIAL PRIMARY KEY,
  tenant_id INTEGER NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  version INTEGER NOT NULL DEFAULT 1,
  exported_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  payload JSONB NOT NULL,
  note VARCHAR(512)
);

CREATE INDEX IF NOT EXISTS idx_erp_backups_tenant_exported
  ON erp_backups (tenant_id, exported_at DESC);

-- Exchange rates (existing API)
CREATE TABLE IF NOT EXISTS exchange_rates (
  id SERIAL PRIMARY KEY,
  tenant_id INTEGER NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  from_currency VARCHAR(16) NOT NULL,
  to_currency VARCHAR(16) NOT NULL DEFAULT 'HKD',
  rate NUMERIC(18, 6) NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE (tenant_id, from_currency, to_currency)
);

-- Legacy normalized tables (AR import / future use)
CREATE TABLE IF NOT EXISTS customers (
  id SERIAL PRIMARY KEY,
  tenant_id INTEGER NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  name VARCHAR(255) NOT NULL,
  type VARCHAR(64) DEFAULT 'CUSTOMER'
);

CREATE TABLE IF NOT EXISTS ar_invoices (
  id SERIAL PRIMARY KEY,
  tenant_id INTEGER NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  customer_id INTEGER REFERENCES customers(id),
  cust_po VARCHAR(128),
  po_completed BOOLEAN DEFAULT FALSE,
  invoice_number VARCHAR(128),
  currency VARCHAR(16),
  invoice_amount NUMERIC(18, 2),
  amount_hkd NUMERIC(18, 2),
  invoice_date TIMESTAMPTZ,
  due_date TIMESTAMPTZ,
  payment_received_date TIMESTAMPTZ,
  po_type VARCHAR(64),
  remark TEXT,
  balance NUMERIC(18, 2),
  status VARCHAR(64)
);

-- Row Level Security (optional — enable when using Supabase Auth)
ALTER TABLE erp_backups ENABLE ROW LEVEL SECURITY;
ALTER TABLE exchange_rates ENABLE ROW LEVEL SECURITY;

-- Service role bypasses RLS; anon blocked until policies added
COMMENT ON TABLE erp_backups IS 'Full ERP JSON snapshots from Settings → Cloud Sync';
