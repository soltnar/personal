import assert from 'node:assert/strict';
import {summarizeAttendance} from '../supabase/functions/personnel-report/engine.mjs';
import {kitchenMessage,reportMessage,subscriptionAssignments,moscowDay} from '../supabase/functions/personnel-notify/message.mjs';
const date='2026-10-26',venue='Белинского, 61 Самурай',centre='Коминтерна 166 CALL CENTRE',other='Ленина, 36 Самурай';
const event=(person,action,time,position='Повар',department='Самурай Белинского 61')=>({person,actionType:action,dateTime:`${date} ${time}`,position,department});
const days=[{work_date:date,status:'ready',rows:[event('Без выхода',1,'09:00:00'),event('Без входа',0,'20:00:00'),event('Нормальная смена',1,'09:00:00'),event('Нормальная смена',0,'20:00:00'),event('Бармен секрет',1,'09:00:00','Бармен'),event('Администратор зала',1,'09:00:00','Администратор'),event('Оператор сети',1,'09:00:00','Оператор',centre),event('Другая точка',1,'09:00:00','Повар','Самурай Ленина 36'),event('Неизвестная должность',1,'09:00:00','')]}];
const sales=[{sale_date:date,restaurant:'Самурай, Белинского, д.61 (КУХНЯ)',revenue:100000},{sale_date:date,restaurant:'Самурай, Белинского, д.61 (БАР)',revenue:20000}];
const metrics=[{dateIso:date,restaurant:venue,orders:10,revenue:30000}],network=[{dateIso:date,orders:500,revenue:800000}];
const rows=summarizeAttendance(days,sales,metrics,network,[date]),message=kitchenMessage(rows,venue,date);
assert(message.includes('Без выхода — нет отметки выхода'));
assert(message.includes('Без входа — нет отметки входа'));
assert(message.includes('Смены: 1'));
assert(message.includes('100 000 ₽/смену'));
for(const name of ['Бармен секрет','Администратор зала','Другая точка','Неизвестная должность'])assert(!message.includes(name));
assert(!message.includes('Нормальная смена —'));
assert.throws(()=>kitchenMessage(rows,'Другая точка',date));
assert(kitchenMessage(summarizeAttendance(days,[]),venue,date).includes('Сводка неполная'));
const bar=reportMessage(rows,venue,date,['bar']);assert(bar.includes('Бармен секрет'));assert(bar.includes('20 000 ₽'));assert(!bar.includes('Без выхода'));assert(!bar.includes('Администратор зала'));
const hall=reportMessage(rows,venue,date,['hall']);assert(hall.includes('Администратор зала'));assert(hall.includes('90 000 ₽'));assert(!hall.includes('Бармен секрет'));
const combined=reportMessage(rows,venue,date,['hall','bar']);assert(combined.includes('Администратор зала'));assert(combined.includes('Бармен секрет'));assert(!combined.includes('Без выхода'));
const operators=reportMessage(rows,centre,date,['operators']);assert(operators.includes('Оператор сети'));assert(operators.includes('800 000 ₽'));assert(operators.includes('Заказы: 500'));assert(!operators.includes('Бармен секрет'));
assert.throws(()=>subscriptionAssignments([{restaurant:venue,roles:['operators']}]));
assert.throws(()=>subscriptionAssignments([{restaurant:venue,roles:['admin']}]));
assert.throws(()=>subscriptionAssignments([{restaurant:venue,roles:[]}]));
assert.throws(()=>subscriptionAssignments([{restaurant:venue,roles:['bar']},{restaurant:venue,roles:['hall']}]));
assert.equal(moscowDay(new Date('2026-10-26T21:01:00Z')).yesterday,date);
assert.equal(moscowDay(new Date('2026-10-27T07:00:00Z')).hour,10);
let handler,token='',isOwner=true,ready=true,records=[],recipients=[],sent=[],lastSave,historyReady=true;
globalThis.Deno={env:{get:key=>key==='TELEGRAM_BOT_TOKEN'?token:key==='SUPABASE_URL'?'https://test.invalid':'key'},serve:fn=>handler=fn};
globalThis.fetch=async(url,init={})=>{
 const parsed=new URL(url),path=parsed.pathname;
 if(parsed.hostname==='viruchki.ilgp.ru')return Response.json(parsed.searchParams.get('action')==='catalog'?{restaurants:[],warehouses:[]}:{users:[{id:123456,firstName:'Сотрудник',status:'approved',salesAccess:{mode:'all'}}]});
 if(parsed.hostname==='history.ilgp.ru')return historyReady?Response.json({metrics,network,completeDates:[date]}):Response.json({error:'missing'},{status:503});
 if(parsed.hostname==='api.telegram.org'){sent.push(JSON.parse(init.body));return Response.json({ok:true,result:{message_id:77}});}
 if(path==='/auth/v1/user')return Response.json({email:isOwner?'soltnar@gmail.com':'other@example.org',email_confirmed_at:date});
 if(path.endsWith('personnel_telegram_config'))return Response.json([]);
 if(path.endsWith('attendance_daily_cache'))return Response.json(ready?days:[{...days[0],status:'loading'}]);
 if(path.endsWith('revenue_entries'))return Response.json(sales);
 if(path.endsWith('personnel_telegram_recipients')){if(init.method){lastSave=JSON.parse(init.body);return Response.json([lastSave]);}return Response.json(recipients);}
 if(path.endsWith('personnel_telegram_log')){
  if(init.method==='POST'){
   const body=JSON.parse(init.body);if(records.some(row=>row.restaurant===body.restaurant))return Response.json({}, {status:409});
   const record={...body,id:String(records.length)};records.push(record);return Response.json([record]);
  }
  if(init.method==='PATCH'){Object.assign(records.find(row=>row.id===parsed.searchParams.get('id').slice(3)),JSON.parse(init.body));return Response.json([]);}
  const venueFilter=parsed.searchParams.get('restaurant');return Response.json(venueFilter?records.filter(row=>row.restaurant===venueFilter.slice(3)):records);
 }
 throw new Error('Unexpected URL '+path);
};
await import('../supabase/functions/personnel-notify/index.ts');
const call=(action,body,auth=true)=>handler(new Request(`https://test.invalid/?action=${action}`,{method:body?'POST':'GET',headers:auth?{authorization:'Bearer test'}:{},...(body?{body:JSON.stringify(body)}:{})}));
assert.equal((await call('settings',null,false)).status,401);
isOwner=false;for(const action of ['settings','directory','save','preview','send'])assert.equal((await call(action,action==='settings'?undefined:{})).status,403);isOwner=true;
assert.equal((await (await call('settings')).json()).configured,false);
assert.equal((await (await call('directory')).json()).users[0].id,'123456');
const assignments=[{restaurant:venue,roles:['hall','bar']},{restaurant:centre,roles:['operators']}];
assert.equal((await call('save',{assignments,name:'Руководитель',chat_id:'123456',send_hour:10,enabled:false})).status,200);
assert.equal(lastSave.assignments.length,2);assert.deepEqual(lastSave.assignments[0].roles,['hall','bar']);
assert.equal((await call('save',{assignments:[{restaurant:venue,roles:['operators']}],name:'Руководитель',chat_id:'123456',enabled:false})).status,400);
let preview=await (await call('preview',{date,assignments})).json();assert.equal(preview.messages.length,2);assert(preview.messages[1].text.includes('Оператор сети'));
historyReady=false;preview=await (await call('preview',{date,assignments:[{restaurant:venue,roles:['kitchen']},{restaurant:centre,roles:['operators']}]})).json();assert(preview.messages[0].text);assert(preview.messages[1].error);historyReady=true;
ready=false;preview=await (await call('preview',{date,assignments})).json();assert(preview.messages.every(item=>item.error));ready=true;
token='test';const id='11111111-1111-1111-1111-111111111111';recipients=[{id,assignments,chat_id:'123456',enabled:true}];
await call('send',{date,id});await call('send',{date,id});
assert.equal(sent.length,2,'One message per enterprise, repeated calls do not send duplicates');
assert(sent[0].text.includes('Бармен секрет'));assert(!sent[0].text.includes('Без выхода'));assert(sent[1].text.includes('Оператор сети'));
assert(records.every(row=>row.status==='sent'));assert.deepEqual(records[0].roles,['hall','bar']);
// Missing history must not leak or substitute the entire enterprise revenue.
records=[];sent=[];historyReady=false;
await call('send',{date,id});assert.equal(sent.length,0);assert.equal(records.length,0);
historyReady=true;records=[];sent=[];
sales.push({sale_date:date,restaurant:'Самурай, Ленина, д.36 (КУХНЯ)',revenue:50000},{sale_date:date,restaurant:'Самурай, Коминтерна, д.115 (КУХНЯ)',revenue:40000});
recipients[0].assignments=[...assignments,{restaurant:other,roles:['kitchen']},{restaurant:'Коминтерна, 115 Самурай',roles:['kitchen']}];
const firstBatch=await (await call('send',{date,id})).json();assert.equal(firstBatch.hasMore,true);assert.equal(sent.length,3);
await call('send',{date,id});assert.equal(sent.length,4,'Interrupted batches resume remaining venues only');
await call('send',{date,id});assert.equal(sent.length,4);
console.log('PASS: user scopes, kitchen/hall/bar/operators formulas and names, multi-venue messages, history completeness, owner-only settings, validation, Moscow time and duplicate protection');
