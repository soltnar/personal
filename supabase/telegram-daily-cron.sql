-- Replace the legacy single-recipient job with the shared subscription scheduler.
-- Recipients, enabled flags, Moscow hours and scopes are read by personnel-notify.
-- One request sends at most three messages; subsequent minute ticks drain the queue.
-- Existing daily claims prevent duplicate messages. Vault references are retained.
do $$
declare
  previous record;
  shared_command text;
begin
  select * into strict previous from cron.job
    where jobname = 'personnel-telegram-shalaev';
  if position('&recipient_id=4b074297-3ac4-4be8-853a-e354ca09f298' in previous.command)=0 then
    raise exception 'Unexpected legacy job scope';
  end if;
  shared_command := replace(previous.command,
    '&recipient_id=4b074297-3ac4-4be8-853a-e354ca09f298', '');
  perform cron.schedule('personnel-telegram-daily', '* 4-17 * * *', shared_command);
  perform cron.unschedule(previous.jobid);
end $$;
