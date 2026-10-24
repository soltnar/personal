import assert from 'node:assert/strict';
import {createHmac} from 'node:crypto';
import {readHistoryMetrics} from '../supabase/functions/personnel-report/history.mjs';
import {summarizeAttendance} from '../supabase/functions/personnel-report/engine.mjs';
const day='2026-10-23';
const venues=[['Самурай Белинского 61 — Доставка','Белинского, 61 Самурай',129,253726.44],['Самурай Верхне-Печерская 14Б — Доставка','Верхне-Печерская, 14Б Самурай',119,226639.59],['Самурай Волжская 13 — Доставка','Волжская, 13 Самурай',185,347541.13],['Самурай Гагарина 224а — Доставка','Кварц, Гагарина, 224А',142,272122.69],['Самурай Каспарус — Доставка','Циолковского, 19А Самурай',132,244123.86],['Самурай Коминтерна 166 — Доставка','Коминтерна 166 CALL CENTRE',190,359636.15],['Самурай Ленина 36 — Доставка','Ленина, 36 Самурай',158,266308.33],['Самурай Моторный пер. 2к1 — Доставка','Моторный, 2/1 доставка Самурай',230,447227.96]];
const payload={completeDates:[day],metrics:venues.map(([restaurant,,orders,revenue])=>({restaurant,dateIso:day,orders,revenue}))};
const history=await readHistoryMetrics(day,day,()=> 'test-secret',async(url,init)=>{
 const path=new URL(url).pathname+new URL(url).search,stamp=init.headers['X-Personnel-Time'];
 assert.equal(init.headers['X-Personnel-Signature'],createHmac('sha256','test-secret').update(`personnel-delivery\n${stamp}\n${path}`).digest('hex'));
 return Response.json(payload);
});
const rows=summarizeAttendance([],[],history.metrics);
assert.equal(rows.length,8);
for(const [,restaurant,orders,revenue] of venues){const row=rows.find(row=>row.restaurant===restaurant);assert(row,restaurant);assert.equal(row.deliveryOrders,orders);assert.equal(row.deliveryRevenue,revenue);}
await assert.rejects(readHistoryMetrics(day,day,()=> 'secret',async()=>Response.json({metrics:[],completeDates:[]})),/Неполный/);
console.log('PASS: signed history bridge, real eight-venue totals and canonical mapping, incomplete days rejected');
