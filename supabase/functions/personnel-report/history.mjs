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
  if(response.status===409&&Array.isArray(payload.missingDates)){
    const error=new Error('Неполный период истории заказов');
    error.missingDates=payload.missingDates;
    throw error;
  }
  if(!response.ok||!Array.isArray(payload.metrics)||!Array.isArray(payload.completeDates))throw new Error('Не загружены дни истории заказов');
  for(let date=from;date<=to;date=new Date(Date.parse(date)+86400000).toISOString().slice(0,10))if(!payload.completeDates.includes(date))throw new Error('Неполный период истории заказов');
  for(const metric of payload.metrics)if(!payload.completeDates.includes(metric.dateIso)||typeof metric.restaurant!=='string'||!Number.isInteger(metric.orders)||metric.orders<0||(metric.revenue!==null&&!Number.isFinite(metric.revenue)))throw new Error('Некорректные данные истории заказов');
  if(!Array.isArray(payload.network))throw new Error('Нет суммы заказов сети');
  for(const day of payload.network)if(!payload.completeDates.includes(day.dateIso)||!Number.isInteger(day.orders)||day.orders<0||(day.revenue!==null&&!Number.isFinite(day.revenue)))throw new Error('Некорректные итоги сети');
  if(payload.completeDates.some(date=>!payload.network.some(day=>day.dateIso===date)))throw new Error('Неполные итоги сети');
  return payload;
}


export async function readPayrollDelivery(from,to,env,request=fetch){
  const dates=[];
  for(let date=from;date<=to;date=new Date(Date.parse(date)+86400000).toISOString().slice(0,10)) dates.push(date);
  let reports;
  try {
    reports=[await readHistoryMetrics(from,to,env,request)];
  } catch(error) {
    if(!Array.isArray(error?.missingDates)) throw error;
    const missing=new Set(error.missingDates);
    if(!dates.some(date=>missing.has(date))) throw error;
    const ranges=[];
    for(const [index,date] of dates.entries()){
      if(missing.has(date)) continue;
      if(index&& !missing.has(dates[index-1])) ranges.at(-1)[1]=date;
      else ranges.push([date,date]);
    }
    reports=await Promise.all(ranges.map(([start,end])=>readHistoryMetrics(start,end,env,request)));
  }
  const values=new Map(reports.flatMap(report=>report.completeDates.map(date=>[date,0])));
  for(const report of reports) for(const metric of report.metrics){
    const amount=values.get(metric.dateIso);
    values.set(metric.dateIso,amount===null||metric.revenue===null?null:amount+metric.revenue);
  }
  const covered=dates.filter(date=>Number.isFinite(values.get(date)));
  const observedRevenue=covered.length?Math.round(covered.reduce((sum,date)=>sum+values.get(date),0)*100)/100:null;
  return {observedRevenue,revenue:covered.length===dates.length?observedRevenue:null,revenueDays:covered.length,expectedDays:dates.length,
    missingRevenueDays:dates.filter(date=>!covered.includes(date)).map(date=>Number(date.slice(-2)))};
}
