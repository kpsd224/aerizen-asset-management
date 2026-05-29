-- Aerizen Asset Management v3.0 Enterprise Workflow
-- Jalankan di Supabase SQL Editor.

create table if not exists public.aerizen_records (
  id text primary key,
  collection text not null,
  payload jsonb not null,
  updated_at timestamptz default now(),
  deleted_at timestamptz
);

create index if not exists aerizen_records_collection_idx on public.aerizen_records(collection);
create index if not exists aerizen_records_updated_at_idx on public.aerizen_records(updated_at desc);

alter table public.aerizen_records enable row level security;

-- Untuk demo/internal app. Sesuaikan policy ketika production.
drop policy if exists "Aerizen read all" on public.aerizen_records;
drop policy if exists "Aerizen write all" on public.aerizen_records;

create policy "Aerizen read all"
on public.aerizen_records
for select
using (true);

create policy "Aerizen write all"
on public.aerizen_records
for all
using (true)
with check (true);
