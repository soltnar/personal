import assert from 'node:assert/strict';
import {isDeepStrictEqual} from 'node:util';
import fs from 'node:fs';
import vm from 'node:vm';
import {stripTypeScriptTypes} from 'node:module';
import {nextAutomaticDate,failureValues,syncActive,retryDue} from '../supabase/functions/personnel-api/schedule.mjs';
let now=Date.parse('2026-11-09T10:00:00Z');
class Clock extends Date {constructor(...args){super(...(args.length?args:[now]));}static now(){return now;}}
let handler,background=[],failPage=1,calls=[],late=false,deny=false;
const oldTime='2026-11-06T10:00:00Z';
const days=new Map(['2026-11-05','2026-11-06'].map(day=>[day,{work_date:day,status:'ready',rows:[{person:'Old snapshot',cacheVersion:4}],updated_at:oldTime,error:null,sync_progress:null}]));
const fetchMock=async(url,init={})=>{
 assert(String(url).length<1500,'Queue claims must not serialize raw pages into URLs');
 const u=new URL(url);
 if(u.pathname==='/auth/v1/user')return Response.json({email:deny?'other@example.com':'soltnar@gmail.com'});
 if(u.pathname.endsWith('/attendance_sync_config'))return Response.json([{}]);
 if(u.hostname==='online.sbis.ru')return Response.json({token:'token'});
 if(u.hostname==='online.saby.ru'){
   const request=JSON.parse(init.body),page=request.params['Навигация'].d[2],day=request.params['Фильтр'].d[7][0];calls.push([day,page]);
   if(page===failPage)throw new Error('Simulated timeout');
   return Response.json({result:{s:['Source','Person','DateTime','ActionType','LastName'].map(n=>({n})),d:page<5?Array.from({length:30},(_,i)=>[3,page*30+i,day+' '+(late?'20:00:00':'09:00:00'),late?0:1,'Person '+page+' '+i]):[],n:page<5}});
 }
 if(u.pathname.endsWith('/attendance_daily_cache')){
   const date=u.searchParams.get('work_date')?.replace('eq.','');
   if(init.method==='PATCH'){
     const entry=days.get(date);if(!entry)return Response.json([]);
     const equal=u.searchParams.get('sync_progress');
     if(equal==='is.null'&&entry.sync_progress!==null)return Response.json([]);
     if(equal?.startsWith('eq.')&&!isDeepStrictEqual(entry.sync_progress,JSON.parse(equal.slice(3))))return Response.json([]);
     for(const key of ['requested_at','lease_id','lease_at']){
       const filter=u.searchParams.get('sync_progress->>'+key);
       if(filter==='is.null'&&entry.sync_progress?.[key]!=null)return Response.json([]);
       if(filter?.startsWith('eq.')&&entry.sync_progress?.[key]!==filter.slice(3))return Response.json([]);
     }
     Object.assign(entry,JSON.parse(init.body));return Response.json([entry]);
   }
   if(init.method==='POST'){
     const entry=JSON.parse(init.body);if(days.has(entry.work_date))return Response.json([]);days.set(entry.work_date,entry);return Response.json([entry]);
   }
   return Response.json([...days.values()].filter(e=>{
     for(const filter of u.searchParams.getAll('work_date')){
       const [op,value]=[filter.slice(0,filter.indexOf('.')),filter.slice(filter.indexOf('.')+1)];
       if(op==='gte'&&e.work_date<value||op==='lte'&&e.work_date>value)return false;
     }return true;
   }));
 }
 throw new Error('Unexpected request '+url);
};
const source=fs.readFileSync(new URL('../supabase/functions/personnel-api/index.ts',import.meta.url),'utf8').replace(/^import .*;\n/gm,'');
const context=vm.createContext({nextAutomaticDate,failureValues,syncActive,retryDue,fetch:fetchMock,Deno:{env:{get:()=> 'https://test.invalid'},serve:fn=>handler=fn},EdgeRuntime:{waitUntil:p=>background.push(p)},console:{log(){},error(){}},Request,Response,URL,Intl,Date:Clock,Map,Set,Promise,TextEncoder,crypto,AbortSignal,setTimeout:fn=>{fn();return 0;},clearTimeout});
vm.runInContext(stripTypeScriptTypes(source),context);
const before=new Clock().toISOString();
const poll=async(query=`from=2026-11-05&to=2026-11-06&refresh=force&before=${before}`)=>{
 const response=await handler(new Request('https://test.invalid/?'+query,{headers:{Authorization:'Bearer owner','x-cron-key':'test'}}));
 const result=await response.json();for(const task of background.splice(0))await task;return {status:response.status,...result};
};
let result=await poll();assert.equal(result.status,200);
assert([...days.values()].every(e=>e.sync_progress?.requested_at),'Entire period queued durably before source work');
let first=days.get('2026-11-05');
assert.equal(first.status,'ready');assert.equal(first.rows[0].person,'Old snapshot');assert.equal(first.updated_at,oldTime);
assert.equal(first.sync_progress.next_page,1,'Saved page survives failure of ready-day refresh');
assert.equal(first.sync_progress.rows.length,30);assert(first.sync_progress.retry_after);
assert.equal(first.sync_progress.paused,true);
// Another day can make progress while this one waits for its source cooldown.
await poll();assert.equal(days.get('2026-11-06').sync_progress.next_page,1);
failPage=-1;now+=31000;calls=[];
await poll();assert.deepEqual(calls.map(c=>c[1]),[1,2,3,4],'Resume saved page rather than reread page zero');
assert.equal(first.rows[0].person,'Old snapshot','Four-page batch is still not a published snapshot');
await poll();assert.equal(first.sync_progress,null);assert.equal(first.rows.length,150);assert.equal(first.error,null);
// Cron, without any browser force call, continues the second queued day.
await poll('refresh=recent');await poll('refresh=recent');
assert.equal(days.get('2026-11-06').sync_progress,null);assert.equal(days.get('2026-11-06').rows.length,150);
result=await poll();assert.equal(result.complete,true);
// A fresh explicit request starts from page zero and replaces the completed day
// only after the late exits have been retrieved in their entirety.
now+=1000;late=true;calls=[];
const fresh=`from=2026-11-05&to=2026-11-05&refresh=force&before=${new Clock().toISOString()}`;
await poll(fresh);assert.equal(calls[0][1],0);assert(first.rows.every(e=>e.actionType===1));
await poll(fresh);assert(first.rows.every(e=>e.actionType===0));
// A worker whose lease was replaced cannot publish its stale result.
first.sync_progress={requested_at:before,lease_id:'new-worker',paused:true};
const snapshot=structuredClone(first.rows);
await context.syncCacheDates(['2026-11-05'],[{...structuredClone(first),sync_progress:{...first.sync_progress,lease_id:'obsolete'}}]);
assert.deepEqual(first.rows,snapshot);assert.equal(first.sync_progress.lease_id,'new-worker');
deny=true;assert.equal((await poll()).status,403);
console.log('PASS: durable multi-day queue, ready snapshot preserved on timeout, page resume, bounded batch, browser-free cron continuation, late exits, stale-worker isolation, owner authorization');
const recent=day=>({work_date:day,status:'ready',updated_at:new Clock().toISOString()});
const today='2026-11-09';
assert.equal(nextAutomaticDate([recent('2026-11-08'),recent('2026-11-07'),recent('2026-11-06')],today,now),null);
const old={...recent('2026-11-07'),updated_at:new Clock(now-7*3600000).toISOString()};
assert.equal(nextAutomaticDate([recent('2026-11-08'),old,recent('2026-11-06')],today,now),'2026-11-07');
assert.equal(nextAutomaticDate([recent('2026-11-08'),{...old,sync_progress:{requested_at:before,paused:true}}],today,now),'2026-11-07');
assert.equal(syncActive({status:'ready',sync_progress:{lease_at:new Clock().toISOString(),paused:false}},now),true);
assert.equal(syncActive({status:'ready',sync_progress:{lease_at:new Clock(now-121000).toISOString(),paused:false}},now),false);
console.log('PASS: automatic reread of recent complete days, manual queue priority, ready-snapshot leases expire safely');
