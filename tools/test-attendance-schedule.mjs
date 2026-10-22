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
assert.equal(nextAutomaticDate(days,'2026-10-23',now+86400000),'2026-10-11');
assert.equal(nextAutomaticDate(days.filter(e=>e.work_date!=='2026-10-18'),today,now),'2026-10-18');
assert.equal(nextAutomaticDate([],today,now),'2026-10-21');
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
console.log('PASS: ready days unchanged, old holes retried, once per Moscow day, new yesterday, active lock, failure preserves stored events');
