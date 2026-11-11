-- 점심 메뉴 투표 앱(index.html)용 Supabase 테이블
-- 사용법: Supabase 대시보드 → SQL Editor → New query → 이 파일 내용 전체 붙여넣기 → Run
-- 여러 번 실행해도 괜찮아요(이미 있으면 건너뜀).

-- 1) 메뉴 목록
create table if not exists public.lunch_menus (
  id uuid primary key default gen_random_uuid(),
  name text not null check (char_length(btrim(name)) between 1 and 30),
  created_at timestamptz not null default now()
);
-- 대소문자만 다른 같은 메뉴 이름 방지
create unique index if not exists lunch_menus_name_key on public.lunch_menus (lower(name));

-- 2) 투표 (이름 1개당 1표, 메뉴 삭제 시 표도 함께 삭제)
create table if not exists public.lunch_votes (
  voter text primary key check (char_length(btrim(voter)) between 1 and 20),
  menu_id uuid not null references public.lunch_menus(id) on delete cascade,
  updated_at timestamptz not null default now()
);

-- 3) 접근 권한
-- 로그인 없는 팀 앱이라, 공개 키(anon/publishable)로 읽기·추가·수정·삭제를 허용해요.
-- 앱 주소를 아는 사람은 누구나 투표·삭제할 수 있다는 뜻이에요.
alter table public.lunch_menus enable row level security;
alter table public.lunch_votes enable row level security;

grant select, insert, delete on public.lunch_menus to anon;
grant select, insert, update, delete on public.lunch_votes to anon;

drop policy if exists "lunch_menus_select" on public.lunch_menus;
drop policy if exists "lunch_menus_insert" on public.lunch_menus;
drop policy if exists "lunch_menus_delete" on public.lunch_menus;
create policy "lunch_menus_select" on public.lunch_menus for select to anon using (true);
create policy "lunch_menus_insert" on public.lunch_menus for insert to anon with check (true);
create policy "lunch_menus_delete" on public.lunch_menus for delete to anon using (true);

drop policy if exists "lunch_votes_select" on public.lunch_votes;
drop policy if exists "lunch_votes_insert" on public.lunch_votes;
drop policy if exists "lunch_votes_update" on public.lunch_votes;
drop policy if exists "lunch_votes_delete" on public.lunch_votes;
create policy "lunch_votes_select" on public.lunch_votes for select to anon using (true);
create policy "lunch_votes_insert" on public.lunch_votes for insert to anon with check (true);
create policy "lunch_votes_update" on public.lunch_votes for update to anon using (true) with check (true);
create policy "lunch_votes_delete" on public.lunch_votes for delete to anon using (true);
