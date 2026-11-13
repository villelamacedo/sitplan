-- 2026-11-13 - harden tenant_settings and contacts.metadata
--
-- Follow-up to 2026-06-29-create-tenant-settings.sql. That file may already be
-- applied, so the fixes live here instead of editing it. Every statement is
-- idempotent and safe to re-run.

-- 1) Redundant index: UNIQUE (tenant_id, key) already creates an identical
--    btree index, so idx_tenant_settings_tenant_key only costs extra writes
--    and disk.
DROP INDEX IF EXISTS idx_tenant_settings_tenant_key;

-- 2) updated_at must never be NULL and should change on every UPDATE,
--    not only when the application remembers to set it.
UPDATE tenant_settings SET updated_at = now() WHERE updated_at IS NULL;
ALTER TABLE tenant_settings
  ALTER COLUMN updated_at SET NOT NULL;

ALTER TABLE tenant_settings
  ADD COLUMN IF NOT EXISTS created_at TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT now();

CREATE OR REPLACE FUNCTION set_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_tenant_settings_updated_at ON tenant_settings;
CREATE TRIGGER trg_tenant_settings_updated_at
  BEFORE UPDATE ON tenant_settings
  FOR EACH ROW EXECUTE FUNCTION set_updated_at();

-- 3) Reject empty / whitespace-only keys.
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint WHERE conname = 'tenant_settings_key_not_blank'
  ) THEN
    ALTER TABLE tenant_settings
      ADD CONSTRAINT tenant_settings_key_not_blank CHECK (btrim(key) <> '');
  END IF;
END $$;

-- 4) contacts.metadata: the 06-29 migration left it nullable, so code has to
--    handle both NULL and '{}'. Backfill and lock it down to always be a
--    JSON object.
UPDATE contacts SET metadata = '{}'::jsonb WHERE metadata IS NULL;
ALTER TABLE contacts
  ALTER COLUMN metadata SET DEFAULT '{}'::jsonb,
  ALTER COLUMN metadata SET NOT NULL;

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint WHERE conname = 'contacts_metadata_is_object'
  ) THEN
    ALTER TABLE contacts
      ADD CONSTRAINT contacts_metadata_is_object
      CHECK (jsonb_typeof(metadata) = 'object');
  END IF;
END $$;

-- 5) Fast filtering on metadata (e.g. WHERE metadata @> '{"source":"site"}').
--    Note: on a large contacts table this locks writes while it builds. If your
--    migration runner does not wrap files in a transaction, prefer
--    CREATE INDEX CONCURRENTLY.
CREATE INDEX IF NOT EXISTS idx_contacts_metadata_gin
  ON contacts USING GIN (metadata jsonb_path_ops);
