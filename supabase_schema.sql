-- DOD Dermatology 근태·연차·OT 관리 - Supabase 스키마
-- Supabase 프로젝트의 SQL Editor에서 이 파일 내용을 그대로 실행하세요.

create table if not exists employees (
  id text primary key,
  name text not null,
  team text default '미지정',
  position text default '',
  hire_date date,
  opening_leave_minutes integer default 0,
  opening_ot_minutes integer default 0,
  custom_start text default '',
  custom_end text default '',
  active boolean default true,
  created_at timestamptz default now()
);

create table if not exists attendance (
  employee_id text not null,
  date date not null,
  employee_name text,
  dow integer,
  checkin text,
  checkout text,
  updated_at timestamptz default now(),
  primary key (employee_id, date)
);

create index if not exists attendance_date_idx on attendance (date);

create table if not exists ledger (
  id text primary key,
  employee_id text not null,
  type text not null, -- 'leave' | 'ot'
  direction text not null, -- 'use' | 'adjust'
  minutes integer not null,
  date date,
  note text,
  created_at timestamptz default now()
);

create table if not exists upload_log (
  id text primary key,
  file_name text,
  uploaded_at timestamptz default now(),
  months text,
  row_count integer,
  removed integer
);

-- Row Level Security 활성화
alter table employees enable row level security;
alter table attendance enable row level security;
alter table ledger enable row level security;
alter table upload_log enable row level security;

-- 내부 관리 도구용 전체 허용 정책
-- 주의: 이 앱은 로그인 기능이 없고 anon key로만 접근합니다.
-- URL과 anon key를 아는 사람은 누구나 읽기/쓰기가 가능합니다 (기존 공유 방식과 동일한 수준).
drop policy if exists "allow all employees" on employees;
create policy "allow all employees" on employees for all using (true) with check (true);

drop policy if exists "allow all attendance" on attendance;
create policy "allow all attendance" on attendance for all using (true) with check (true);

drop policy if exists "allow all ledger" on ledger;
create policy "allow all ledger" on ledger for all using (true) with check (true);

drop policy if exists "allow all upload_log" on upload_log;
create policy "allow all upload_log" on upload_log for all using (true) with check (true);
