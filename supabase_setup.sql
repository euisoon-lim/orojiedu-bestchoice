-- ============================================================
-- 오로지교육 베스트 초이스 업로드 파이프라인 (Supabase SQL Editor에서 실행)
-- ============================================================
insert into storage.buckets (id, name, public, file_size_limit)
values ('bestchoice-inbox','bestchoice-inbox', false, 52428800)
on conflict (id) do nothing;

create table if not exists public.bestchoice_config (
  key text primary key,
  value text not null,
  updated_at timestamptz not null default now()
);
alter table public.bestchoice_config enable row level security;   -- 정책 없음 → 서버(service role) 전용

create table if not exists public.bestchoice_uploads (
  id bigserial primary key,
  file_name text not null,
  storage_path text not null,
  size_bytes bigint,
  uploaded_at timestamptz not null default now(),
  dispatched boolean not null default false,
  note text
);
alter table public.bestchoice_uploads enable row level security;   -- 정책 없음 → 서버 전용

insert into public.bestchoice_config(key,value) values
 ('passcode','<PASSCODE>'),
 ('gh_pat','<GITHUB_PAT>'),
 ('gh_repo','euisoon-lim/orojiedu-bestchoice')
on conflict (key) do update set value=excluded.value, updated_at=now();

select key, length(value) as len from public.bestchoice_config;
