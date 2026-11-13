-- SITPLAN - acesso de sócio ao Modo TA + CRM básico
-- Execute este arquivo no SQL Editor do Supabase depois do supabase-schema.sql.
-- Objetivo: permitir que outro usuário autenticado veja Modo TA e CRM básico,
-- sem ficha completa, histórico, indicadores, apólices, agenda ou estratégia.

create table if not exists public.sitplan_shares (
  owner_user_id uuid not null references auth.users(id) on delete cascade,
  viewer_user_id uuid not null references auth.users(id) on delete cascade,
  role text not null default 'viewer' check (role in ('viewer')),
  active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  primary key (owner_user_id, viewer_user_id)
);

alter table public.sitplan_shares enable row level security;

drop policy if exists "sitplan_shares_owner_manage" on public.sitplan_shares;
drop policy if exists "sitplan_shares_viewer_read" on public.sitplan_shares;

create policy "sitplan_shares_owner_manage" on public.sitplan_shares
  for all
  using (auth.uid() = owner_user_id)
  with check (auth.uid() = owner_user_id);

create policy "sitplan_shares_viewer_read" on public.sitplan_shares
  for select
  using (auth.uid() = viewer_user_id and active = true);

drop policy if exists "contacts_shared_viewer_select" on public.contacts;
drop policy if exists "activities_shared_viewer_select" on public.activities;
drop policy if exists "weekly_goals_shared_viewer_select" on public.weekly_goals;

-- Importante:
-- Não recriamos políticas de SELECT completo para viewer.
-- O sócio acessa apenas a função sanitizada abaixo.

drop function if exists public.sitplan_ta_viewer_contacts();

create or replace function public.sitplan_ta_viewer_contacts()
returns table (
  id uuid,
  name text,
  phone text,
  whatsapp text,
  stage text,
  priority text,
  follow_up text,
  source text,
  referrer text,
  temperature text,
  last_interaction text,
  ta_date text,
  oi_scheduled_date text,
  next_action text,
  created_at timestamptz
)
language sql
security definer
set search_path = public, auth
as $$
  select
    c.id,
    coalesce(c.data->>'name', '') as name,
    coalesce(c.data->>'phone', '') as phone,
    coalesce(c.data->>'whatsapp', '') as whatsapp,
    coalesce(c.data->>'stage', 'Base') as stage,
    coalesce(c.data->>'priority', 'Media') as priority,
    coalesce(c.data->>'followUp', '') as follow_up,
    coalesce(c.data->>'source', '') as source,
    coalesce(c.data->>'referrer', '') as referrer,
    coalesce(c.data->>'temperature', 'Morno') as temperature,
    coalesce(c.data->>'lastInteraction', '') as last_interaction,
    coalesce(c.data->>'taDate', '') as ta_date,
    coalesce(c.data->>'oiScheduledDate', '') as oi_scheduled_date,
    coalesce(c.data->>'nextAction', '') as next_action,
    c.created_at
  from public.contacts c
  where exists (
    select 1
    from public.sitplan_shares share
    where share.owner_user_id = c.user_id
      and share.viewer_user_id = auth.uid()
      and share.active = true
      and share.role = 'viewer'
  )
    and coalesce(c.data->>'doNotContact', 'false') <> 'true'
  order by c.created_at;
$$;

create or replace function public.sitplan_grant_viewer(viewer_email text)
returns void
language plpgsql
security definer
set search_path = public, auth
as $$
declare
  viewer_id uuid;
begin
  if auth.uid() is null then
    raise exception 'Usuário não autenticado.';
  end if;

  select id into viewer_id
  from auth.users
  where lower(email) = lower(viewer_email)
  limit 1;

  if viewer_id is null then
    raise exception 'Visualizador não encontrado. Crie o usuário no Authentication > Users antes de conceder acesso.';
  end if;

  if viewer_id = auth.uid() then
    raise exception 'Não é necessário compartilhar a base com o próprio usuário.';
  end if;

  insert into public.sitplan_shares (owner_user_id, viewer_user_id, role, active, updated_at)
  values (auth.uid(), viewer_id, 'viewer', true, now())
  on conflict (owner_user_id, viewer_user_id)
  do update set active = true, role = 'viewer', updated_at = now();
end;
$$;

create or replace function public.sitplan_revoke_viewer(viewer_email text)
returns void
language plpgsql
security definer
set search_path = public, auth
as $$
declare
  viewer_id uuid;
begin
  if auth.uid() is null then
    raise exception 'Usuário não autenticado.';
  end if;

  select id into viewer_id
  from auth.users
  where lower(email) = lower(viewer_email)
  limit 1;

  if viewer_id is null then
    return;
  end if;

  update public.sitplan_shares
  set active = false, updated_at = now()
  where owner_user_id = auth.uid()
    and viewer_user_id = viewer_id;
end;
$$;

revoke all on function public.sitplan_grant_viewer(text) from public;
revoke all on function public.sitplan_revoke_viewer(text) from public;
revoke all on function public.sitplan_ta_viewer_contacts() from public;
grant execute on function public.sitplan_grant_viewer(text) to authenticated;
grant execute on function public.sitplan_revoke_viewer(text) to authenticated;
grant execute on function public.sitplan_ta_viewer_contacts() to authenticated;

-- Como conceder acesso:
-- 1. Crie o usuário do sócio em Authentication > Users.
-- 2. Entre no SITPLAN com o seu usuário dono da base.
-- 3. No painel "Sócio operacional", digite o e-mail do sócio e clique em "Liberar TA + CRM".
--
-- Alternativa técnica via API/RPC autenticada:
--    rpc sitplan_grant_viewer('email-do-socio@exemplo.com')
--
-- Como revogar:
-- 1. Entre no SITPLAN com o usuário dono.
-- 2. No painel "Socio operacional", digite o e-mail e clique em "Revogar".
