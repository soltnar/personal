// Automatic loading never refreshes a completed day. Failed attempts wait until
// the next Moscow calendar day; dates older than five days remain in the queue.
export function nextAutomaticDate(entries, today, now = Date.now(), staleMs = 300000) {
  if(entries.some(e=>e.status==='loading' && now-Date.parse(e.updated_at)<staleMs)) return null;
  const byDate=new Map(entries.map(e=>[e.work_date,e]));
  const yesterday=new Date(Date.parse(today+'T12:00:00Z')-86400000).toISOString().slice(0,10);
  const first=entries.length?entries.reduce((date,e)=>e.work_date<date?e.work_date:date,yesterday):yesterday;
  for(let date=first;date<=yesterday;date=new Date(Date.parse(date+'T12:00:00Z')+86400000).toISOString().slice(0,10)) {
    const entry=byDate.get(date);
    if(entry?.status==='ready')continue;
    // UTC+3 is fixed for Moscow. Compare actual date, not a rolling 24h window.
    const attempted=entry?.updated_at ? new Date(Date.parse(entry.updated_at)+10800000).toISOString().slice(0,10) : '';
    if(entry && attempted>=today)continue;
    return date;
  }
  return null;
}

export function failureValues(previous, message, elapsedMs) {
  // Omit rows entirely on failure: never overwrite the last stored events.
  return {status:previous?.status==='ready'?'ready':'error',error:message,elapsed_ms:elapsedMs};
}
