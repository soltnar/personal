-- User-centric subscriptions; existing kitchen-only settings retain their scope.
alter table public.personnel_telegram_recipients add column assignments jsonb;
update public.personnel_telegram_recipients set assignments=jsonb_build_array(jsonb_build_object('restaurant',restaurant,'roles',jsonb_build_array('kitchen')));
alter table public.personnel_telegram_recipients alter column assignments set not null;
alter table public.personnel_telegram_recipients add constraint personnel_telegram_assignments_array check(jsonb_typeof(assignments)='array' and jsonb_array_length(assignments) between 1 and 18);
alter table public.personnel_telegram_recipients add constraint personnel_telegram_recipients_chat_key unique(chat_id);
alter table public.personnel_telegram_log drop constraint personnel_telegram_log_recipient_id_work_date_key;
alter table public.personnel_telegram_log drop constraint personnel_telegram_log_restaurant_chat_id_work_date_key;
alter table public.personnel_telegram_log add constraint personnel_telegram_log_recipient_day_venue_key unique(recipient_id,work_date,restaurant);
alter table public.personnel_telegram_log add constraint personnel_telegram_log_chat_day_venue_key unique(chat_id,work_date,restaurant);
alter table public.personnel_telegram_log add column roles text[] not null default array['kitchen'];
