-- 2026-11-13 - SITPLAN CRM (Supabase): indexes + trustworthy updated_at
--
-- Run in the Supabase SQL Editor after supabase-schema.sql and
-- supabase-viewer-access.sql. Additive and idempotent: no data is changed,
-- no column is renamed, the web app needs no update.

-- 1) Indexes for the RLS filter (auth.uid() = user_id).
--    Every query the app makes goes through RLS policies that filter on
--    user_id, but only the primary key is indexed, so each load/poll
--    (every 15 s per open tab) scans the whole table.
create index if not exists contacts_user_id_idx
  on public.contacts (user_id);
create index if not exists activities_user_date_idx
  on public.activities (user_id, activity_date, activity_time);
create index if not exists activities_contact_id_idx
  on public.activities (contact_id);
create index if not exists sitplan_shares_viewer_idx
  on public.sitplan_shares (viewer_user_id) where active;

-- 2) updated_at = when the record really changed.
--    The app saves a full snapshot and sends updated_at = now() for EVERY row
--    on every save, so today updated_at only means "last save of anything".
--    This trigger keeps the old timestamp when the content did not change and
--    uses the server clock (not the device clock) when it did. That makes it
--    usable for "recently edited" lists, audits and future conflict detection.
create or replace function public.sitplan_touch_updated_at()
returns trigger
language plpgsql
set search_path = public
as $$
begin
  if tg_op = 'UPDATE'
     and (to_jsonb(new) - 'updated_at') = (to_jsonb(old) - 'updated_at') then
    new.updated_at := old.updated_at;
  else
    new.updated_at := now();
  end if;
  return new;
end;
$$;

drop trigger if exists contacts_touch_updated_at on public.contacts;
create trigger contacts_touch_updated_at
  before insert or update on public.contacts
  for each row execute function public.sitplan_touch_updated_at();

drop trigger if exists activities_touch_updated_at on public.activities;
create trigger activities_touch_updated_at
  before insert or update on public.activities
  for each row execute function public.sitplan_touch_updated_at();

drop trigger if exists weekly_goals_touch_updated_at on public.weekly_goals;
create trigger weekly_goals_touch_updated_at
  before insert or update on public.weekly_goals
  for each row execute function public.sitplan_touch_updated_at();

drop trigger if exists sitplan_shares_touch_updated_at on public.sitplan_shares;
create trigger sitplan_shares_touch_updated_at
  before insert or update on public.sitplan_shares
  for each row execute function public.sitplan_touch_updated_at();
