import assert from 'node:assert/strict';
import {summarizeAttendance} from '../supabase/functions/personnel-report/engine.mjs';
import {kitchenMessage,moscowDay} from '../supabase/functions/personnel-notify/message.mjs';
const date='2026-10-26',venue='Белинского, 61 Самурай';
const event=(person,action,time,position='Повар',department='Самурай Белинского 61')=>({person,actionType:action,dateTime:`${date} ${time}`,position,department});
const days=[{work_date:date,status:'ready',rows:[event('Без выхода',1,'09:00:00'),event('Без входа',0,'20:00:00'),event('Нормальная смена',1,'09:00:00'),event('Нормальная смена',0,'20:00:00'),event('Бармен секрет',1,'09:00:00','Бармен'),event('Другая точка',1,'09:00:00','Повар','Самурай Ленина 36'),event('Неизвестная должность',1,'09:00:00','')]}];
const sales=[{sale_date:date,restaurant:'Самурай, Белинского, д.61 (КУХНЯ)',revenue:100000}];
const rows=summarizeAttendance(days,sales),message=kitchenMessage(rows,venue,date);
assert(message.includes('Без выхода — нет отметки выхода'));
assert(message.includes('Без входа — нет отметки входа'));
assert(message.includes('Смены кухни: 1'));
assert(message.includes('100 000 ₽/смену'));
for(const name of ['Бармен секрет','Другая точка','Неизвестная должность'])assert(!message.includes(name));
assert(!message.includes('Нормальная смена —'));
assert.throws(()=>kitchenMessage(rows,'Другая точка',date));
assert(kitchenMessage(summarizeAttendance(days,[]),venue,date).includes('Сводка неполная'));
assert.equal(moscowDay(new Date('2026-10-26T21:01:00Z')).yesterday,'2026-10-26');
assert.equal(moscowDay(new Date('2026-10-27T06:00:00Z')).hour,9);
let handler,token='',isOwner=true,ready=true,records=[],recipients=[],sent=[];
globalThis.Deno={env:{get:key=>key==='TELEGRAM_BOT_TOKEN'?token:key==='SUPABASE_URL'?'https://test.invalid':'key'},serve:fn=>handler=fn};
globalThis.fetch=async(url,init={})=>{
  const parsed=new URL(url),path=parsed.pathname;
  if(parsed.hostname==='api.telegram.org'){sent.push(JSON.parse(init.body));return Response.json({ok:true,result:{message_id:77}});}
  if(path==='/auth/v1/user')return Response.json({email:isOwner?'soltnar@gmail.com':'other@example.org',email_confirmed_at:date});
  if(path.endsWith('personnel_telegram_config'))return Response.json([]);
  if(path.endsWith('attendance_daily_cache'))return Response.json(ready?days:[{...days[0],status:'loading'}]);
  if(path.endsWith('revenue_entries'))return Response.json(sales);
  if(path.endsWith('personnel_telegram_recipients'))return Response.json(recipients);
  if(path.endsWith('personnel_telegram_log')){
    if(init.method==='POST'){
      if(records.length)return Response.json({}, {status:409});
      records.push({...JSON.parse(init.body),id:'log'});return Response.json(records);
    }
    if(init.method==='PATCH'){Object.assign(records[0],JSON.parse(init.body));return Response.json(records);}
    return Response.json(records);
  }
  throw new Error('Unexpected URL '+path);
};
await import('../supabase/functions/personnel-notify/index.ts');
const call=(action,body,auth=true)=>handler(new Request(`https://test.invalid/?action=${action}`,{method:body?'POST':'GET',headers:auth?{authorization:'Bearer test'}:{},...(body?{body:JSON.stringify(body)}:{})}));
assert.equal((await call('settings',null,false)).status,401);
isOwner=false;assert.equal((await call('settings')).status,403);isOwner=true;
assert.equal((await (await call('settings')).json()).configured,false);
assert.equal((await call('preview',{date,restaurant:venue})).status,200);
ready=false;assert.equal((await call('preview',{date,restaurant:venue})).status,500);ready=true;
assert.equal((await call('save',{restaurant:venue,name:'Шеф',chat_id:'@public',send_hour:9,enabled:true})).status,400);
token='test';const id='11111111-1111-1111-1111-111111111111';recipients=[{id,restaurant:venue,chat_id:'123456',enabled:true}];
await call('send',{date,id});await call('send',{date,id});
assert.equal(sent.length,1,'Repeated calls do not send the daily report twice');
assert.equal(sent[0].chat_id,'123456');assert.equal(records[0].status,'sent');
assert(!sent[0].text.includes('Бармен секрет'));
console.log('PASS: names and missing marks including zero shifts, existing shift totals, kitchen/venue isolation, missing data, Moscow dates, owner auth, send once');
