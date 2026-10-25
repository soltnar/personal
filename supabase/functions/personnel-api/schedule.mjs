// Completed days stay intact. Paused batches resume, transient errors retry
// after ten minutes, and old missing days remain in the queue.
export function nextAutomaticDate(entries, today, now = Date.now(), staleMs = 300000) {
  if(entries.some(e=>e.status==='loading' && !e.sync_progress?.paused && now-Date.parse(e.updated_at)<staleMs)) return null;
  const byDate=new Map(entries.map(e=>[e.work_date,e]));
  const yesterday=new Date(Date.parse(today+'T12:00:00Z')-86400000).toISOString().slice(0,10);
  const first=entries.length?entries.reduce((date,e)=>e.work_date<date?e.work_date:date,yesterday):yesterday;
  const candidates=[yesterday];
  for(let date=first;date<yesterday;date=new Date(Date.parse(date+'T12:00:00Z')+86400000).toISOString().slice(0,10))candidates.push(date);
  // Serve the fresh daily report before spending the remaining slots on old holes.
  for(const date of candidates) {
    const entry=byDate.get(date);
    if(entry?.status==='ready')continue;
    // Retry transient failures during the same day; paused batches resume immediately.
    if(entry && !entry.sync_progress?.paused && now-Date.parse(entry.updated_at)<10*60000)continue;
    return date;
  }
  return null;
}

export function failureValues(previous, message, elapsedMs) {
  // Omit rows entirely on failure: never overwrite the last stored events.
  return {status:previous?.status==='ready'?'ready':'error',error:message,elapsed_ms:elapsedMs};
}
