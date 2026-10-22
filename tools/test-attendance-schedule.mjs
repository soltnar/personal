import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { stripTypeScriptTypes } from 'node:module';
import { nextAutomaticDate, failureValues } from '../supabase/functions/personnel-api/schedule.mjs';
const today='2026-10-22', now=Date.parse('2026-10-22T03:00:00Z');
const ready=day=>({work_date:day,status:'ready',updated_at:'2026-10-20T00:00:00Z'});
const days=Array.from({length:11},(_,i)=>ready(`2026-10-${11+i}`));
assert.equal(nextAutomaticDate(days,today,now),null,'Completed days never refreshed');
days[0]={...days[0],status:'error'};
assert.equal(nextAutomaticDate(days,today,now),'2026-10-11','Old error retried beyond five days');
days[0].updated_at='2026-10-21T22:00:00Z';
assert.equal(nextAutomaticDate(days,today,now),null,'Failed today in Moscow waits until tomorrow');
assert.equal(nextAutomaticDate(days,'2026-10-23',now+86400000),'2026-10-22','New yesterday has priority');
assert.equal(nextAutomaticDate([...days,ready('2026-10-22')],'2026-10-23',now+86400000),'2026-10-11','Old errors retry next day after the daily report');
assert.equal(nextAutomaticDate(days.filter(e=>e.work_date!=='2026-10-18'),today,now),'2026-10-18');
assert.equal(nextAutomaticDate([],today,now),'2026-10-21');
assert.equal(nextAutomaticDate([{work_date:'2026-06-15',status:'error',updated_at:'2026-10-20T00:00:00Z'}],today,now),'2026-10-21','New yesterday takes priority over historical backfill');
assert.equal(nextAutomaticDate([...days,{work_date:'2026-10-10',status:'loading',updated_at:new Date(now-1000).toISOString()}],today,now),null);
assert(!Object.hasOwn(failureValues(ready('2026-10-18'),'timeout',50000),'rows'));
assert.equal(failureValues(ready('2026-10-18'),'timeout',50000).status,'ready');
assert.equal(failureValues(null,'timeout',50000).status,'error');
// Execute the actual handler with a failed Saby request, preserving a prior day.
let handler,background;
let saved={...ready('2026-10-18'),rows:[{cacheVersion:4,person:'Test',dateTime:'2026-10-18 10:00:00'}],error:null};
const originalDate=saved.updated_at;
const fetchMock=async(url,init={})=>{
  const parsed=new URL(url);
  if(parsed.pathname==='/auth/v1/user')return Response.json({email:'soltnar@gmail.com'});
  if(parsed.hostname==='online.sbis.ru')throw new Error('Simulated timeout');
  if(parsed.pathname.endsWith('/attendance_daily_cache')){
    if(init.method==='POST'||init.method==='PATCH'){Object.assign(saved,JSON.parse(init.body));return Response.json([saved]);}
    return Response.json([saved]);
  }
  throw new Error('Unexpected request '+url);
};
const source=fs.readFileSync(new URL('../supabase/functions/personnel-api/index.ts',import.meta.url),'utf8').replace(/^import .*;\n/gm,'');
const context=vm.createContext({nextAutomaticDate,failureValues,fetch:fetchMock,Deno:{env:{get:()=> 'https://test.invalid'},serve:fn=>{handler=fn;}},EdgeRuntime:{waitUntil:promise=>{background=promise;}},console:{log(){},error(){}},Request,Response,URL,Intl,Date,Map,Set,Promise,TextEncoder,crypto,AbortSignal,setTimeout,clearTimeout});
vm.runInContext(stripTypeScriptTypes(source),context);
const response=await handler(new Request('https://test.invalid/?from=2026-10-18&to=2026-10-18&refresh=force&before=2026-10-22T00:00:00Z',{headers:{Authorization:'Bearer test'}}));
assert.equal(response.status,200);await background;
assert.equal(saved.status,'ready');assert.equal(saved.rows.length,1);assert.equal(saved.updated_at,originalDate);assert(saved.error);
assert(saved.sync_progress.failed_at);
background=null;
const poll=await handler(new Request('https://test.invalid/?from=2026-10-18&to=2026-10-18&refresh=force&before=2026-10-22T00:00:00Z',{headers:{Authorization:'Bearer test'}}));
assert.equal(background,null,'A failed forced refresh of a previously ready day must not restart on polling');
assert.equal((await poll.json()).errors.length,1,'Forced refresh exposes its failure while preserving the ready snapshot');
console.log('PASS: ready days unchanged, old holes retried, once per Moscow day, new yesterday, active lock, failure preserves stored events');
let calls=[],fail=true,checkpoint;
context.setTimeout=fn=>{fn();return 0;};
context.fetch=async(_url,init)=>{
  const request=JSON.parse(init.body),page=request.params['Навигация'].d[2];calls.push(page);
  if(fail&&page===1)throw new Error('Intermittent source failure');
  return Response.json({result:{s:['Source','DateTime','ActionType','LastName','FirstName','Person'].map(n=>({n})),d:page<2?Array.from({length:30},(_,i)=>[3,'2026-10-19 10:00:00',0,'Test',`Employee ${page}-${i}`,page*30+i]):[],n:page<2}});
};
await assert.rejects(context.fetchAttendance('2026-10-19','2026-10-19','token',null,async value=>{checkpoint=structuredClone(value);}));
assert.equal(checkpoint.next_page,1);assert.equal(checkpoint.rows.length,30);
assert.deepEqual(calls,[0,1,1]);calls=[];fail=false;
const resumed=await context.fetchAttendance('2026-10-19','2026-10-19','token',checkpoint,async value=>{checkpoint=structuredClone(value);});
assert.deepEqual(calls,[1,2]);assert.equal(resumed.rows.length,60);assert.equal(checkpoint.complete,true);
calls=[];context.fetch=async()=>{calls.push(1);return new Response('',{status:429,headers:{'Retry-After':'60'}});};
await assert.rejects(context.fetchAttendance('2026-10-19','2026-10-19','token'),/429/);assert.equal(calls.length,1);
console.log('PASS: interrupted pagination resumes without losing or duplicating rows; HTTP 429 does not trigger immediate retries');
