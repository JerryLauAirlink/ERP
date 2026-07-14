-- Fix integer overflow: frontend record IDs use millisecond timestamps (>2^31).
-- Run in Supabase Dashboard → SQL Editor after 002_live_sync_entities.sql

ALTER TABLE erp_entity_records
  ALTER COLUMN entity_id TYPE BIGINT USING entity_id::BIGINT;

ALTER TABLE erp_entity_records
  ALTER COLUMN updated_by TYPE BIGINT USING updated_by::BIGINT;
