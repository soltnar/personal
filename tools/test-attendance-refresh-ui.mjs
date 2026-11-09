import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const source=fs.readFileSync(new URL('../script.js',import.meta.url),'utf8');
const start=source.indexOf('async function loadSabyFromApi('),end=source.indexOf('\nasync function loadRevenueFromDatabase(',start);
async function check(persistent,finish){
 let now=Date.parse('2026-11-09T10:00:00Z'),calls=0,reported=0,statuses=[];
 class Clock extends Date{constructor(...args){super(...(args.length?args:[now]));}static now(){return now;}}
 const button={disabled:false,classList:{add(){},remove(){}},setAttribute(){},removeAttribute(){}};
 const context=vm.createContext({Date:Clock,Promise,Array,Math,Error,
 sabyApiLoading:false,supabaseClient:{auth:{getSession:async()=>({data:{session:{access_token:'test'}}})}},
 sabyFromInput:{value:'2026-11-05'},sabyToInput:{value:'2026-11-06'},document:{getElementById:()=>button},loadSabyBtn:button,
 SABY_CACHE_WAIT_MS:20000,SABY_CACHE_POLL_MS:4000,updateRevenueDbButtons(){},setSabyApiStatus:(message,kind)=>statuses.push([message,kind]),
 fetchSabyChunk:async()=>{calls++;if(calls===2){const e=new Error('Temporary connection failure');e.retryable=true;throw e;}return {rows:[{}],readyDates:finish&&calls>=3?['2026-11-05','2026-11-06']:[],pendingDates:finish&&calls>=3?[]:['2026-11-05','2026-11-06'],complete:finish&&calls>=3,diagnostics:{persistentQueue:persistent,progressPage:4,progressRows:120}};},
 setTimeout:fn=>{now+=4000;fn();},prepareSabyData:()=>({attendance:[{}],staff:[]}),applyStaffData(){},applyAttendanceData(){},mappedRecords:[],calculate:()=>[],renderTable(){},loadScopedReport:async()=>reported++});
 vm.runInContext(source.slice(start,end),context);
 await context.loadSabyFromApi({force:true});
 assert.equal(context.sabyApiLoading,false);assert.equal(button.disabled,false);
 return {calls,reported,statuses};
}
let result=await check(true,true);assert.equal(result.reported,1);assert.equal(result.statuses.at(-1)[1],'success');assert(result.calls>=3);
result=await check(true,false);assert.match(result.statuses.at(-1)[0],/продолжается в фоне/);assert.equal(result.statuses.at(-1)[1],'loading');
result=await check(false,false);assert.equal(result.statuses.at(-1)[1],'error');assert.match(result.statuses.at(-1)[0],/Не удалось подтвердить/);
console.log('PASS: refresh UI recovers polling connection loss, completes, distinguishes durable background work from an unconfirmed queue, restores buttons');
