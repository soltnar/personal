-- Applied through Supabase apply_migration: personnel_telegram_delivery.
create table public.personnel_telegram_recipients (
 id uuid primary key default gen_random_uuid(),
 restaurant text not null,
 name text not null,
 chat_id text not null check(chat_id ~ '^-?[0-9]{5,20}$'),
 enabled boolean not null default false,
 send_hour integer not null default 10 check(send_hour between 7 and 18),
 updated_at timestamptz not null default now(),
 unique(restaurant,chat_id)
);
create table public.personnel_telegram_log (
 id uuid primary key default gen_random_uuid(),
 recipient_id uuid not null references public.personnel_telegram_recipients(id),
 restaurant text not null,
 chat_id text not null,
 work_date date not null,
 status text not null check(status in ('sending','sent','review')),
 message_id text,
 error text,
 created_at timestamptz not null default now(),
 updated_at timestamptz not null default now(),
 unique(recipient_id,work_date),
 unique(restaurant,chat_id,work_date)
);
alter table public.personnel_telegram_recipients enable row level security;
alter table public.personnel_telegram_log enable row level security;
revoke all on public.personnel_telegram_recipients,public.personnel_telegram_log from anon,authenticated;
grant all on public.personnel_telegram_recipients,public.personnel_telegram_log to service_role;

-- Applied through Supabase apply_migration: personnel_telegram_private_config.
-- Reuses the revenue bot without exposing its token to browser clients.
create table public.personnel_telegram_config (
 id integer primary key check(id=1),
 bot_token text not null,
 updated_at timestamptz not null default now()
);
alter table public.personnel_telegram_config enable row level security;
revoke all on public.personnel_telegram_config from anon,authenticated;
grant all on public.personnel_telegram_config to service_role;
