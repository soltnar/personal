-- Existing authenticated worker; keep its Vault headers and URL intact.
-- Every minute, including outside the old morning window, resume queued pages.
-- New automatic daily loads start after 05:00 MSK (see schedule.mjs).
select cron.schedule(jobname, '* * * * *', command)
from cron.job
where jobname = 'refresh-attendance-last-5-days';
