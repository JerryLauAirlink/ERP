-- ERP Live Sync — per-entity records for multi-user real-time sync
-- Run in Supabase Dashboard → SQL Editor after 001_initial_schema.sql

CREATE TABLE IF NOT EXISTS erp_entity_records (
  id BIGSERIAL PRIMARY KEY,
  tenant_id INTEGER NOT NULL REFERENCES tenants(id) ON DELETE CASCADE,
  entity_type VARCHAR(64) NOT NULL,
  entity_id BIGINT NOT NULL,
  region VARCHAR(16),
  payload JSONB NOT NULL DEFAULT '{}'::jsonb,
  is_deleted BOOLEAN NOT NULL DEFAULT FALSE,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_by INTEGER,
  UNIQUE (tenant_id, entity_type, entity_id)
);

CREATE INDEX IF NOT EXISTS idx_erp_entity_tenant_type_updated
  ON erp_entity_records (tenant_id, updated_at DESC);

CREATE INDEX IF NOT EXISTS idx_erp_entity_tenant_type
  ON erp_entity_records (tenant_id, entity_type);

CREATE TABLE IF NOT EXISTS erp_sync_meta (
  tenant_id INTEGER PRIMARY KEY REFERENCES tenants(id) ON DELETE CASCADE,
  server_version BIGINT NOT NULL DEFAULT 0,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

INSERT INTO erp_sync_meta (tenant_id, server_version, updated_at)
VALUES (1, 0, NOW())
ON CONFLICT (tenant_id) DO NOTHING;

ALTER TABLE erp_entity_records ENABLE ROW LEVEL SECURITY;
ALTER TABLE erp_sync_meta ENABLE ROW LEVEL SECURITY;

COMMENT ON TABLE erp_entity_records IS 'Per-entity live sync (clients, jobs, AR, AP, etc.)';
COMMENT ON TABLE erp_sync_meta IS 'Monotonic version counter for change polling';
