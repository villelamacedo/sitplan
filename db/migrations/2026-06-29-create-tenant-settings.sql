-- 2026-06-29 - create tenant_settings table and add metadata to contacts

CREATE TABLE IF NOT EXISTS tenant_settings (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  tenant_id UUID NOT NULL,
  key TEXT NOT NULL,
  value TEXT NOT NULL,
  updated_by UUID,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT now(),
  UNIQUE (tenant_id, key)
);

-- Ensure contacts table has a JSONB metadata column for audit/meta
ALTER TABLE contacts
  ADD COLUMN IF NOT EXISTS metadata JSONB DEFAULT '{}';

-- Optional index for faster lookup by tenant and key
CREATE INDEX IF NOT EXISTS idx_tenant_settings_tenant_key ON tenant_settings (tenant_id, key);
