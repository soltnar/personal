begin;

create table if not exists public.personnel_payroll_months (
  month text primary key check (month ~ '^20[0-9]{2}-(0[1-9]|1[0-2])$'),
  groups jsonb not null check (jsonb_typeof(groups) = 'array'),
  updated_at timestamptz not null default now()
);

alter table public.personnel_payroll_months enable row level security;
revoke all on table public.personnel_payroll_months from public, anon, authenticated;
grant select, insert, update on table public.personnel_payroll_months to service_role;

create table if not exists public.personnel_payroll_access (
  identity text primary key check (identity ~ '^(google|saby):.+$')
);

alter table public.personnel_payroll_access enable row level security;
revoke all on table public.personnel_payroll_access from public, anon, authenticated;
grant select on table public.personnel_payroll_access to service_role;

commit;
