// Queue metadata lives beside the last complete snapshot, never in its rows.
export function syncActive(entry, now = Date.now(), staleMs = 120000) {
  const progress=entry?.sync_progress;
  if(progress?.lease_at)return !progress.paused && now-Date.parse(progress.lease_at)<staleMs;
  return entry?.status==='loading' && !progress?.paused && now-Date.parse(entry.updated_at)<300000;
}
export function retryDue(entry, now = Date.now()) {
  return !entry?.sync_progress?.retry_after || Date.parse(entry.sync_progress.retry_after)<=now;
}
export function nextAutomaticDate(entries, today, now = Date.now()) {
  if(entries.some(entry=>syncActive(entry,now)))return null;
  const queued=entries.filter(e=>e.sync_progress?.requested_at&&retryDue(e,now)).sort((a,b)=>a.sync_progress.requested_at.localeCompare(b.sync_progress.requested_at)||a.work_date.localeCompare(b.work_date));
  if(queued.length)return queued[0].work_date;
  if(new Date(now+3*3600000).getUTCHours()<5)return null;
  const byDate=new Map(entries.map(e=>[e.work_date,e]));
  const add=(date,days)=>new Date(Date.parse(date+'T12:00:00Z')+days*86400000).toISOString().slice(0,10);
  const yesterday=add(today,-1);
  const first=entries.reduce((date,e)=>e.work_date<date?e.work_date:date,yesterday);
  const candidates=[yesterday];
  for(let date=first;date<yesterday;date=add(date,1))candidates.push(date);
  for(const date of candidates){
    const entry=byDate.get(date);
    if(entry?.status==='ready')continue;
    if(!retryDue(entry,now))continue;
    if(entry&&!entry.sync_progress?.paused&&now-Date.parse(entry.updated_at)<600000)continue;
    return date;
  }
  // Devices may upload yesterday's exits hours or days after their entries.
  // Re-read complete recent days from page zero, at most once per Moscow calendar day.
  for(let i=1;i<=3;i++){
    const date=add(today,-i),entry=byDate.get(date);
    if(entry?.status==='ready'&&!entry.sync_progress?.requested_at&&retryDue(entry,now)&&new Date(Date.parse(entry.updated_at)+3*3600000).toISOString().slice(0,10)<today)return date;
  }
  return null;
}
export function failureValues(previous,message,elapsedMs){
  return {status:previous?.status==='ready'?'ready':'error',error:message,elapsed_ms:elapsedMs};
}
