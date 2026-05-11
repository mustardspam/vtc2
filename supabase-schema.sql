-- VTC Scorecard Supabase schema
-- Run this in the Supabase SQL editor.
--
-- This replaces the prototype "one JSON row" model with normalized tables,
-- RLS, master data, imports, field feedback, scoring, snapshots, and activity.
-- The scorecard_states table is kept at the bottom as a temporary compatibility
-- bridge for the current static prototype while the frontend is migrated.

create extension if not exists pgcrypto;

do $$ begin
  create type public.app_role as enum ('admin', 'manager', 'viewer', 'builder');
exception when duplicate_object then null;
end $$;

do $$ begin
  create type public.upload_status as enum ('uploaded', 'mapped', 'validated', 'approved', 'rejected', 'archived');
exception when duplicate_object then null;
end $$;

do $$ begin
  create type public.source_data_type as enum (
    'schedule',
    'safety',
    'rework',
    'field_feedback',
    'vendor_master',
    'community_master',
    'supporting'
  );
exception when duplicate_object then null;
end $$;

do $$ begin
  create type public.feedback_status as enum ('needs_review', 'approved', 'rejected');
exception when duplicate_object then null;
end $$;

do $$ begin
  create type public.feedback_category as enum ('kudos', 'complaint');
exception when duplicate_object then null;
end $$;

do $$ begin
  create type public.complaint_severity as enum ('minor', 'major', 'critical');
exception when duplicate_object then null;
end $$;

create or replace function public.touch_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  email text not null,
  full_name text,
  role public.app_role not null default 'viewer',
  can_upload boolean not null default false,
  can_adjust_weights boolean not null default false,
  can_save_snapshots boolean not null default false,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create or replace function public.current_profile_role()
returns public.app_role
language sql
stable
security definer
set search_path = public
as $$
  select role
  from public.profiles
  where id = auth.uid()
    and is_active = true
  limit 1
$$;

create or replace function public.is_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select coalesce(public.current_profile_role() = 'admin'::public.app_role, false)
$$;

create or replace function public.is_manager_or_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select coalesce(public.current_profile_role() in ('admin'::public.app_role, 'manager'::public.app_role), false)
$$;

create or replace function public.can_upload()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select coalesce(
    exists (
      select 1
      from public.profiles
      where id = auth.uid()
        and is_active = true
        and (role = 'admin'::public.app_role or can_upload = true)
    ),
    false
  )
$$;

create or replace function public.can_adjust_weights()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select coalesce(
    exists (
      select 1
      from public.profiles
      where id = auth.uid()
        and is_active = true
        and (role = 'admin'::public.app_role or can_adjust_weights = true)
    ),
    false
  )
$$;

create or replace function public.can_save_snapshots()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select coalesce(
    exists (
      select 1
      from public.profiles
      where id = auth.uid()
        and is_active = true
        and (role = 'admin'::public.app_role or can_save_snapshots = true)
    ),
    false
  )
$$;

create table if not exists public.trade_categories (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  description text,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.vendors (
  id uuid primary key default gen_random_uuid(),
  external_vendor_id text unique,
  name text not null,
  normalized_name text not null unique,
  is_active boolean not null default true,
  created_by uuid references public.profiles(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.trades (
  id uuid primary key default gen_random_uuid(),
  vendor_id uuid not null references public.vendors(id) on delete cascade,
  category_id uuid references public.trade_categories(id),
  name text not null,
  normalized_name text not null,
  is_active boolean not null default true,
  created_by uuid references public.profiles(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (vendor_id, normalized_name)
);

create table if not exists public.communities (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  normalized_name text not null unique,
  is_active boolean not null default true,
  created_by uuid references public.profiles(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.vendor_aliases (
  id uuid primary key default gen_random_uuid(),
  vendor_id uuid not null references public.vendors(id) on delete cascade,
  alias text not null unique,
  normalized_alias text not null unique,
  created_by uuid references public.profiles(id),
  created_at timestamptz not null default now()
);

create table if not exists public.community_aliases (
  id uuid primary key default gen_random_uuid(),
  community_id uuid not null references public.communities(id) on delete cascade,
  alias text not null unique,
  normalized_alias text not null unique,
  created_by uuid references public.profiles(id),
  created_at timestamptz not null default now()
);

create table if not exists public.uploaded_files (
  id uuid primary key default gen_random_uuid(),
  storage_bucket text not null default 'source-files',
  storage_path text not null,
  original_filename text not null,
  content_type text,
  file_size_bytes bigint,
  data_type public.source_data_type not null,
  checksum text,
  uploaded_by uuid references public.profiles(id),
  uploaded_at timestamptz not null default now()
);

create table if not exists public.import_batches (
  id uuid primary key default gen_random_uuid(),
  data_type public.source_data_type not null,
  status public.upload_status not null default 'uploaded',
  source_file_ids uuid[] not null default '{}',
  mapping_json jsonb not null default '{}'::jsonb,
  validation_summary jsonb not null default '{}'::jsonb,
  approved_by uuid references public.profiles(id),
  approved_at timestamptz,
  created_by uuid references public.profiles(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.raw_import_rows (
  id uuid primary key default gen_random_uuid(),
  import_batch_id uuid not null references public.import_batches(id) on delete cascade,
  uploaded_file_id uuid references public.uploaded_files(id),
  row_number int not null,
  raw_data jsonb not null,
  validation_errors jsonb not null default '[]'::jsonb,
  normalized_record_id uuid,
  created_at timestamptz not null default now(),
  unique(import_batch_id, row_number)
);

create table if not exists public.normalized_schedule_records (
  id uuid primary key default gen_random_uuid(),
  import_batch_id uuid references public.import_batches(id),
  raw_import_row_id uuid references public.raw_import_rows(id),
  vendor_id uuid references public.vendors(id),
  trade_id uuid references public.trades(id),
  score_month date not null,
  monthly_jobs int not null default 0,
  no_show_count int not null default 0,
  score_basis jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  check (monthly_jobs >= 0),
  check (no_show_count >= 0),
  check (no_show_count <= monthly_jobs)
);

create table if not exists public.normalized_safety_records (
  id uuid primary key default gen_random_uuid(),
  import_batch_id uuid references public.import_batches(id),
  raw_import_row_id uuid references public.raw_import_rows(id),
  vendor_id uuid references public.vendors(id),
  trade_id uuid references public.trades(id),
  incident_date date,
  score_month date not null,
  category text,
  incident_severity text,
  severity_score numeric not null default 0,
  osha_recordable boolean,
  notes text,
  created_at timestamptz not null default now()
);

create table if not exists public.normalized_rework_records (
  id uuid primary key default gen_random_uuid(),
  import_batch_id uuid references public.import_batches(id),
  raw_import_row_id uuid references public.raw_import_rows(id),
  vendor_id uuid references public.vendors(id),
  trade_id uuid references public.trades(id),
  rework_date date,
  score_month date not null,
  units_affected numeric,
  rework_cost numeric,
  severity text,
  notes text,
  penalty_points numeric not null default 0,
  created_at timestamptz not null default now()
);

create table if not exists public.field_feedback_rules (
  id uuid primary key default gen_random_uuid(),
  category public.feedback_category not null,
  severity public.complaint_severity,
  points numeric not null,
  is_active boolean not null default true,
  updated_by uuid references public.profiles(id),
  updated_at timestamptz not null default now(),
  unique(category, severity)
);

create table if not exists public.field_feedback_records (
  id uuid primary key default gen_random_uuid(),
  source_type text not null default 'widget',
  status public.feedback_status not null default 'needs_review',
  uploaded_file_id uuid references public.uploaded_files(id),
  import_batch_id uuid references public.import_batches(id),
  submitted_by uuid references public.profiles(id),
  submitted_name text,
  submitted_at timestamptz not null default now(),
  vendor_id uuid references public.vendors(id),
  trade_id uuid references public.trades(id),
  community_id uuid references public.communities(id),
  address text,
  lot text,
  category public.feedback_category not null,
  severity public.complaint_severity,
  assigned_points numeric,
  notes text,
  evidence_text text,
  reviewed_by uuid references public.profiles(id),
  reviewed_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  check (
    (category = 'kudos'::public.feedback_category and severity is null)
    or
    (category = 'complaint'::public.feedback_category and severity is not null)
  )
);

create table if not exists public.feedback_attachments (
  id uuid primary key default gen_random_uuid(),
  feedback_id uuid not null references public.field_feedback_records(id) on delete cascade,
  storage_bucket text not null default 'feedback-attachments',
  storage_path text not null,
  original_filename text not null,
  content_type text,
  file_size_bytes bigint,
  uploaded_by uuid references public.profiles(id),
  created_at timestamptz not null default now()
);

create table if not exists public.score_weights (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  safety_weight numeric not null default 25,
  schedule_weight numeric not null default 25,
  rework_weight numeric not null default 12.5,
  field_feedback_weight numeric not null default 37.5,
  is_default boolean not null default false,
  created_by uuid references public.profiles(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  check (safety_weight + schedule_weight + rework_weight + field_feedback_weight = 100)
);

create unique index if not exists one_default_score_weight
on public.score_weights(is_default)
where is_default = true;

create table if not exists public.score_results (
  id uuid primary key default gen_random_uuid(),
  vendor_id uuid not null references public.vendors(id),
  trade_id uuid references public.trades(id),
  score_month date not null,
  safety_score numeric,
  schedule_score numeric,
  rework_score numeric,
  field_feedback_score numeric,
  weighted_total_score numeric,
  rank int,
  supporting_counts jsonb not null default '{}'::jsonb,
  calculation_detail jsonb not null default '{}'::jsonb,
  weight_id uuid references public.score_weights(id),
  calculated_at timestamptz not null default now(),
  unique(vendor_id, trade_id, score_month, weight_id)
);

create table if not exists public.score_snapshots (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  notes text,
  score_month date not null,
  source_file_ids uuid[] not null default '{}',
  weight_config jsonb not null,
  is_locked boolean not null default true,
  created_by uuid references public.profiles(id),
  created_at timestamptz not null default now(),
  deleted_at timestamptz,
  deleted_by uuid references public.profiles(id)
);

create table if not exists public.snapshot_score_results (
  id uuid primary key default gen_random_uuid(),
  snapshot_id uuid not null references public.score_snapshots(id) on delete cascade,
  vendor_id uuid not null,
  trade_id uuid,
  score_month date not null,
  safety_score numeric,
  schedule_score numeric,
  rework_score numeric,
  field_feedback_score numeric,
  weighted_total_score numeric,
  rank int,
  supporting_counts jsonb not null default '{}'::jsonb,
  calculation_detail jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create table if not exists public.activity_log (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references public.profiles(id),
  action_type text not null,
  description text not null,
  uploaded_file_id uuid references public.uploaded_files(id),
  import_batch_id uuid references public.import_batches(id),
  snapshot_id uuid references public.score_snapshots(id),
  before_values jsonb,
  after_values jsonb,
  metadata jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create or replace function public.calculate_monthly_scores(p_score_month date, p_weight_id uuid default null)
returns table (
  vendor_id uuid,
  trade_id uuid,
  score_month date,
  safety_score numeric,
  schedule_score numeric,
  rework_score numeric,
  field_feedback_score numeric,
  weighted_total_score numeric,
  supporting_counts jsonb,
  calculation_detail jsonb
)
language sql
stable
as $$
  with selected_weights as (
    select *
    from public.score_weights
    where id = coalesce(
      p_weight_id,
      (select id from public.score_weights where is_default = true limit 1)
    )
  ),
  vendor_trade as (
    select v.id as vendor_id, t.id as trade_id
    from public.vendors v
    left join public.trades t on t.vendor_id = v.id and t.is_active = true
    where v.is_active = true
  ),
  schedule as (
    select
      vendor_id,
      trade_id,
      sum(monthly_jobs)::numeric as monthly_jobs,
      sum(no_show_count)::numeric as no_show_count,
      case
        when sum(monthly_jobs) > 0
          then greatest(0, least(100, ((sum(monthly_jobs) - sum(no_show_count))::numeric / sum(monthly_jobs)) * 100))
        else null
      end as score
    from public.normalized_schedule_records
    where date_trunc('month', score_month)::date = date_trunc('month', p_score_month)::date
    group by vendor_id, trade_id
  ),
  safety as (
    select
      vendor_id,
      trade_id,
      count(*) as incident_count,
      sum(severity_score) as severity_points,
      greatest(0, 100 - (sum(severity_score) * 10)) as score
    from public.normalized_safety_records
    where date_trunc('month', score_month)::date = date_trunc('month', p_score_month)::date
    group by vendor_id, trade_id
  ),
  rework as (
    select
      vendor_id,
      trade_id,
      count(*) as rework_count,
      sum(penalty_points) as penalty_points,
      greatest(0, 100 - (sum(penalty_points) * 5)) as score
    from public.normalized_rework_records
    where date_trunc('month', score_month)::date = date_trunc('month', p_score_month)::date
    group by vendor_id, trade_id
  ),
  feedback as (
    select
      vendor_id,
      trade_id,
      count(*) as feedback_count,
      avg(assigned_points) as score
    from public.field_feedback_records
    where status = 'approved'::public.feedback_status
      and date_trunc('month', submitted_at)::date = date_trunc('month', p_score_month)::date
    group by vendor_id, trade_id
  ),
  components as (
    select
      vt.vendor_id,
      vt.trade_id,
      date_trunc('month', p_score_month)::date as score_month,
      coalesce(safety.score, 100) as safety_score,
      coalesce(schedule.score, 100) as schedule_score,
      rework.score as rework_score,
      feedback.score as field_feedback_score,
      jsonb_build_object(
        'monthlyJobs', coalesce(schedule.monthly_jobs, 0),
        'noShows', coalesce(schedule.no_show_count, 0),
        'safetyIncidents', coalesce(safety.incident_count, 0),
        'reworkItems', coalesce(rework.rework_count, 0),
        'feedbackItems', coalesce(feedback.feedback_count, 0)
      ) as supporting_counts,
      jsonb_build_object(
        'scheduleFormula', '(monthly_jobs - no_show_count) / monthly_jobs * 100',
        'noShowOnlySchedulePenalty', true,
        'safetyFormula', '100 - sum(severity_score) * 10',
        'reworkFormula', '100 - sum(penalty_points) * 5',
        'fieldFeedbackFormula', 'average approved feedback points'
      ) as calculation_detail
    from vendor_trade vt
    left join schedule on schedule.vendor_id = vt.vendor_id and schedule.trade_id is not distinct from vt.trade_id
    left join safety on safety.vendor_id = vt.vendor_id and safety.trade_id is not distinct from vt.trade_id
    left join rework on rework.vendor_id = vt.vendor_id and rework.trade_id is not distinct from vt.trade_id
    left join feedback on feedback.vendor_id = vt.vendor_id and feedback.trade_id is not distinct from vt.trade_id
  )
  select
    c.vendor_id,
    c.trade_id,
    c.score_month,
    c.safety_score,
    c.schedule_score,
    c.rework_score,
    c.field_feedback_score,
    (
      coalesce(c.safety_score * sw.safety_weight, 0)
      + coalesce(c.schedule_score * sw.schedule_weight, 0)
      + coalesce(c.rework_score * sw.rework_weight, 0)
      + coalesce(c.field_feedback_score * sw.field_feedback_weight, 0)
    )
    / nullif(
      (case when c.safety_score is null then 0 else sw.safety_weight end)
      + (case when c.schedule_score is null then 0 else sw.schedule_weight end)
      + (case when c.rework_score is null then 0 else sw.rework_weight end)
      + (case when c.field_feedback_score is null then 0 else sw.field_feedback_weight end),
      0
    ) as weighted_total_score,
    c.supporting_counts,
    c.calculation_detail || jsonb_build_object(
      'weights',
      jsonb_build_object(
        'safety', sw.safety_weight,
        'schedule', sw.schedule_weight,
        'rework', sw.rework_weight,
        'fieldFeedback', sw.field_feedback_weight
      )
    ) as calculation_detail
  from components c
  cross join selected_weights sw;
$$;

create index if not exists idx_trades_vendor_id on public.trades(vendor_id);
create index if not exists idx_uploaded_files_data_type on public.uploaded_files(data_type, uploaded_at desc);
create index if not exists idx_import_batches_status on public.import_batches(data_type, status, created_at desc);
create index if not exists idx_raw_import_rows_batch on public.raw_import_rows(import_batch_id);
create index if not exists idx_schedule_month_vendor_trade on public.normalized_schedule_records(score_month, vendor_id, trade_id);
create index if not exists idx_safety_month_vendor_trade on public.normalized_safety_records(score_month, vendor_id, trade_id);
create index if not exists idx_rework_month_vendor_trade on public.normalized_rework_records(score_month, vendor_id, trade_id);
create index if not exists idx_feedback_status_month_vendor_trade on public.field_feedback_records(status, submitted_at, vendor_id, trade_id);
create index if not exists idx_score_results_month_total on public.score_results(score_month, weighted_total_score desc);
create index if not exists idx_snapshots_month on public.score_snapshots(score_month, created_at desc);
create index if not exists idx_activity_log_created on public.activity_log(created_at desc);

drop trigger if exists touch_profiles_updated_at on public.profiles;
create trigger touch_profiles_updated_at
before update on public.profiles
for each row execute function public.touch_updated_at();

drop trigger if exists touch_trade_categories_updated_at on public.trade_categories;
create trigger touch_trade_categories_updated_at
before update on public.trade_categories
for each row execute function public.touch_updated_at();

drop trigger if exists touch_vendors_updated_at on public.vendors;
create trigger touch_vendors_updated_at
before update on public.vendors
for each row execute function public.touch_updated_at();

drop trigger if exists touch_trades_updated_at on public.trades;
create trigger touch_trades_updated_at
before update on public.trades
for each row execute function public.touch_updated_at();

drop trigger if exists touch_communities_updated_at on public.communities;
create trigger touch_communities_updated_at
before update on public.communities
for each row execute function public.touch_updated_at();

drop trigger if exists touch_import_batches_updated_at on public.import_batches;
create trigger touch_import_batches_updated_at
before update on public.import_batches
for each row execute function public.touch_updated_at();

drop trigger if exists touch_feedback_records_updated_at on public.field_feedback_records;
create trigger touch_feedback_records_updated_at
before update on public.field_feedback_records
for each row execute function public.touch_updated_at();

drop trigger if exists touch_score_weights_updated_at on public.score_weights;
create trigger touch_score_weights_updated_at
before update on public.score_weights
for each row execute function public.touch_updated_at();

insert into public.score_weights (
  name,
  safety_weight,
  schedule_weight,
  rework_weight,
  field_feedback_weight,
  is_default
)
values ('Workbook default', 25, 25, 12.5, 37.5, true)
on conflict do nothing;

insert into public.field_feedback_rules (category, severity, points)
select 'kudos'::public.feedback_category, null::public.complaint_severity, 100
where not exists (
  select 1 from public.field_feedback_rules
  where category = 'kudos'::public.feedback_category and severity is null
);

insert into public.field_feedback_rules (category, severity, points)
values
  ('complaint', 'minor', 85),
  ('complaint', 'major', 70),
  ('complaint', 'critical', 50)
on conflict (category, severity) do update
set points = excluded.points,
    updated_at = now();

alter table public.profiles enable row level security;
alter table public.trade_categories enable row level security;
alter table public.vendors enable row level security;
alter table public.trades enable row level security;
alter table public.communities enable row level security;
alter table public.vendor_aliases enable row level security;
alter table public.community_aliases enable row level security;
alter table public.uploaded_files enable row level security;
alter table public.import_batches enable row level security;
alter table public.raw_import_rows enable row level security;
alter table public.normalized_schedule_records enable row level security;
alter table public.normalized_safety_records enable row level security;
alter table public.normalized_rework_records enable row level security;
alter table public.field_feedback_rules enable row level security;
alter table public.field_feedback_records enable row level security;
alter table public.feedback_attachments enable row level security;
alter table public.score_weights enable row level security;
alter table public.score_results enable row level security;
alter table public.score_snapshots enable row level security;
alter table public.snapshot_score_results enable row level security;
alter table public.activity_log enable row level security;

drop policy if exists "profiles read own or admin" on public.profiles;
create policy "profiles read own or admin"
on public.profiles
for select
to authenticated
using (id = auth.uid() or public.is_admin());

drop policy if exists "profiles admin write" on public.profiles;
create policy "profiles admin write"
on public.profiles
for all
to authenticated
using (public.is_admin())
with check (public.is_admin());

drop policy if exists "master data read active" on public.trade_categories;
create policy "master data read active"
on public.trade_categories
for select
to anon, authenticated
using (is_active = true or public.is_manager_or_admin());

drop policy if exists "master data write admin" on public.trade_categories;
create policy "master data write admin"
on public.trade_categories
for all
to authenticated
using (public.is_admin())
with check (public.is_admin());

drop policy if exists "vendors read active" on public.vendors;
create policy "vendors read active"
on public.vendors
for select
to anon, authenticated
using (is_active = true or public.is_manager_or_admin());

drop policy if exists "vendors write admin" on public.vendors;
create policy "vendors write admin"
on public.vendors
for all
to authenticated
using (public.is_admin())
with check (public.is_admin());

drop policy if exists "trades read active" on public.trades;
create policy "trades read active"
on public.trades
for select
to anon, authenticated
using (is_active = true or public.is_manager_or_admin());

drop policy if exists "trades write admin" on public.trades;
create policy "trades write admin"
on public.trades
for all
to authenticated
using (public.is_admin())
with check (public.is_admin());

drop policy if exists "communities read active" on public.communities;
create policy "communities read active"
on public.communities
for select
to anon, authenticated
using (is_active = true or public.is_manager_or_admin());

drop policy if exists "communities write admin" on public.communities;
create policy "communities write admin"
on public.communities
for all
to authenticated
using (public.is_admin())
with check (public.is_admin());

drop policy if exists "vendor aliases read" on public.vendor_aliases;
create policy "vendor aliases read"
on public.vendor_aliases
for select
to authenticated
using (true);

drop policy if exists "vendor aliases admin write" on public.vendor_aliases;
create policy "vendor aliases admin write"
on public.vendor_aliases
for all
to authenticated
using (public.is_admin())
with check (public.is_admin());

drop policy if exists "community aliases read" on public.community_aliases;
create policy "community aliases read"
on public.community_aliases
for select
to authenticated
using (true);

drop policy if exists "community aliases admin write" on public.community_aliases;
create policy "community aliases admin write"
on public.community_aliases
for all
to authenticated
using (public.is_admin())
with check (public.is_admin());

drop policy if exists "uploads read managers" on public.uploaded_files;
create policy "uploads read managers"
on public.uploaded_files
for select
to authenticated
using (public.is_manager_or_admin());

drop policy if exists "uploads insert allowed" on public.uploaded_files;
create policy "uploads insert allowed"
on public.uploaded_files
for insert
to authenticated
with check (public.can_upload());

drop policy if exists "imports managers all" on public.import_batches;
create policy "imports managers all"
on public.import_batches
for all
to authenticated
using (public.is_manager_or_admin())
with check (public.is_manager_or_admin());

drop policy if exists "raw rows managers all" on public.raw_import_rows;
create policy "raw rows managers all"
on public.raw_import_rows
for all
to authenticated
using (public.is_manager_or_admin())
with check (public.is_manager_or_admin());

drop policy if exists "schedule read authenticated" on public.normalized_schedule_records;
create policy "schedule read authenticated"
on public.normalized_schedule_records
for select
to authenticated
using (public.current_profile_role() is not null and public.current_profile_role() <> 'builder'::public.app_role);

drop policy if exists "schedule write managers" on public.normalized_schedule_records;
create policy "schedule write managers"
on public.normalized_schedule_records
for all
to authenticated
using (public.is_manager_or_admin())
with check (public.is_manager_or_admin());

drop policy if exists "safety read authenticated" on public.normalized_safety_records;
create policy "safety read authenticated"
on public.normalized_safety_records
for select
to authenticated
using (public.current_profile_role() is not null and public.current_profile_role() <> 'builder'::public.app_role);

drop policy if exists "safety write managers" on public.normalized_safety_records;
create policy "safety write managers"
on public.normalized_safety_records
for all
to authenticated
using (public.is_manager_or_admin())
with check (public.is_manager_or_admin());

drop policy if exists "rework read authenticated" on public.normalized_rework_records;
create policy "rework read authenticated"
on public.normalized_rework_records
for select
to authenticated
using (public.current_profile_role() is not null and public.current_profile_role() <> 'builder'::public.app_role);

drop policy if exists "rework write managers" on public.normalized_rework_records;
create policy "rework write managers"
on public.normalized_rework_records
for all
to authenticated
using (public.is_manager_or_admin())
with check (public.is_manager_or_admin());

drop policy if exists "feedback rules read authenticated" on public.field_feedback_rules;
create policy "feedback rules read authenticated"
on public.field_feedback_rules
for select
to authenticated
using (true);

drop policy if exists "feedback rules admin write" on public.field_feedback_rules;
create policy "feedback rules admin write"
on public.field_feedback_rules
for all
to authenticated
using (public.is_admin())
with check (public.is_admin());

drop policy if exists "feedback read scoped" on public.field_feedback_records;
create policy "feedback read scoped"
on public.field_feedback_records
for select
to authenticated
using (
  public.is_manager_or_admin()
  or (
    public.current_profile_role() = 'viewer'::public.app_role
    and status = 'approved'::public.feedback_status
  )
  or submitted_by = auth.uid()
);

drop policy if exists "builders insert own feedback" on public.field_feedback_records;
create policy "builders insert own feedback"
on public.field_feedback_records
for insert
to authenticated
with check (
  public.current_profile_role() in ('builder'::public.app_role, 'manager'::public.app_role, 'admin'::public.app_role)
  and coalesce(submitted_by, auth.uid()) = auth.uid()
);

drop policy if exists "public widget insert feedback" on public.field_feedback_records;
create policy "public widget insert feedback"
on public.field_feedback_records
for insert
to anon
with check (
  source_type = 'widget'
  and status = 'needs_review'::public.feedback_status
  and submitted_by is null
  and assigned_points is null
  and reviewed_by is null
  and reviewed_at is null
);

drop policy if exists "feedback managers update" on public.field_feedback_records;
create policy "feedback managers update"
on public.field_feedback_records
for update
to authenticated
using (public.is_manager_or_admin())
with check (public.is_manager_or_admin());

drop policy if exists "feedback attachments read scoped" on public.feedback_attachments;
create policy "feedback attachments read scoped"
on public.feedback_attachments
for select
to authenticated
using (
  public.is_manager_or_admin()
  or exists (
    select 1
    from public.field_feedback_records f
    where f.id = feedback_id
      and f.submitted_by = auth.uid()
  )
);

drop policy if exists "feedback attachments insert own" on public.feedback_attachments;
create policy "feedback attachments insert own"
on public.feedback_attachments
for insert
to authenticated
with check (
  public.current_profile_role() in ('builder'::public.app_role, 'manager'::public.app_role, 'admin'::public.app_role)
  and uploaded_by = auth.uid()
);

drop policy if exists "public widget insert attachment metadata" on public.feedback_attachments;
create policy "public widget insert attachment metadata"
on public.feedback_attachments
for insert
to anon
with check (
  uploaded_by is null
  and exists (
    select 1
    from public.field_feedback_records f
    where f.id = feedback_id
      and f.source_type = 'widget'
      and f.status = 'needs_review'::public.feedback_status
      and f.submitted_by is null
  )
);

drop policy if exists "weights read nonbuilders" on public.score_weights;
create policy "weights read nonbuilders"
on public.score_weights
for select
to authenticated
using (public.current_profile_role() is not null and public.current_profile_role() <> 'builder'::public.app_role);

drop policy if exists "weights write allowed" on public.score_weights;
create policy "weights write allowed"
on public.score_weights
for all
to authenticated
using (public.can_adjust_weights())
with check (public.can_adjust_weights());

drop policy if exists "score results read nonbuilders" on public.score_results;
create policy "score results read nonbuilders"
on public.score_results
for select
to authenticated
using (public.current_profile_role() is not null and public.current_profile_role() <> 'builder'::public.app_role);

drop policy if exists "score results managers write" on public.score_results;
create policy "score results managers write"
on public.score_results
for all
to authenticated
using (public.is_manager_or_admin())
with check (public.is_manager_or_admin());

drop policy if exists "snapshots read nonbuilders" on public.score_snapshots;
create policy "snapshots read nonbuilders"
on public.score_snapshots
for select
to authenticated
using (
  deleted_at is null
  and public.current_profile_role() is not null
  and public.current_profile_role() <> 'builder'::public.app_role
);

drop policy if exists "snapshots insert allowed" on public.score_snapshots;
create policy "snapshots insert allowed"
on public.score_snapshots
for insert
to authenticated
with check (public.can_save_snapshots());

drop policy if exists "snapshots admin update" on public.score_snapshots;
create policy "snapshots admin update"
on public.score_snapshots
for update
to authenticated
using (public.is_admin())
with check (public.is_admin());

drop policy if exists "snapshot results read nonbuilders" on public.snapshot_score_results;
create policy "snapshot results read nonbuilders"
on public.snapshot_score_results
for select
to authenticated
using (public.current_profile_role() is not null and public.current_profile_role() <> 'builder'::public.app_role);

drop policy if exists "snapshot results insert allowed" on public.snapshot_score_results;
create policy "snapshot results insert allowed"
on public.snapshot_score_results
for insert
to authenticated
with check (public.can_save_snapshots());

drop policy if exists "activity read managers" on public.activity_log;
create policy "activity read managers"
on public.activity_log
for select
to authenticated
using (public.is_manager_or_admin());

drop policy if exists "activity insert authenticated" on public.activity_log;
create policy "activity insert authenticated"
on public.activity_log
for insert
to authenticated
with check (user_id = auth.uid() or public.is_manager_or_admin());

-- Temporary compatibility table for the current static prototype.
-- Keep this until the frontend writes normalized tables directly.
create table if not exists public.scorecard_states (
  id text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.scorecard_states enable row level security;

drop policy if exists "Allow VTC scorecard read" on public.scorecard_states;
drop policy if exists "Allow VTC scorecard insert" on public.scorecard_states;
drop policy if exists "Allow VTC scorecard update" on public.scorecard_states;

create policy "Allow VTC scorecard read"
on public.scorecard_states
for select
to anon, authenticated
using (id = 'vtc-main');

create policy "Allow VTC scorecard insert"
on public.scorecard_states
for insert
to anon, authenticated
with check (id = 'vtc-main');

create policy "Allow VTC scorecard update"
on public.scorecard_states
for update
to anon, authenticated
using (id = 'vtc-main')
with check (id = 'vtc-main');
