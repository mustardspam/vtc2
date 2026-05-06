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
to anon
using (id = 'vtc-main');

create policy "Allow VTC scorecard insert"
on public.scorecard_states
for insert
to anon
with check (id = 'vtc-main');

create policy "Allow VTC scorecard update"
on public.scorecard_states
for update
to anon
using (id = 'vtc-main')
with check (id = 'vtc-main');
