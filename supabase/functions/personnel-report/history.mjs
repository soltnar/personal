// Read ready daily totals from the same saved reports shown at history.ilgp.ru.
export async function readHistoryMetrics(from,to,env,request=fetch){
  const secret=env('SABY_APP_SECRET');
  if(!secret)throw new Error('Не настроена связь с историей заказов');
  const path=`/api/personnel-delivery?from=${from}&to=${to}`,timestamp=String(Math.floor(Date.now()/1000));
  const encoder=new TextEncoder(),key=await crypto.subtle.importKey('raw',encoder.encode(secret),{name:'HMAC',hash:'SHA-256'},false,['sign']);
  const signed=await crypto.subtle.sign('HMAC',key,encoder.encode(`personnel-delivery\n${timestamp}\n${path}`));
  const signature=Array.from(new Uint8Array(signed),n=>n.toString(16).padStart(2,'0')).join('');
  const response=await request(`https://history.ilgp.ru${path}`,{signal:AbortSignal.timeout(25000),headers:{'X-Personnel-Time':timestamp,'X-Personnel-Signature':signature}});
  const payload=await response.json();
  if(!response.ok||!Array.isArray(payload.metrics)||!Array.isArray(payload.completeDates))throw new Error('Не загружены дни истории заказов');
  for(let date=from;date<=to;date=new Date(Date.parse(date)+86400000).toISOString().slice(0,10))if(!payload.completeDates.includes(date))throw new Error('Неполный период истории заказов');
  for(const metric of payload.metrics)if(!payload.completeDates.includes(metric.dateIso)||typeof metric.restaurant!=='string'||!Number.isInteger(metric.orders)||metric.orders<0||(metric.revenue!==null&&!Number.isFinite(metric.revenue)))throw new Error('Некорректные данные истории заказов');
  return payload;
}
