import { summarizeAttendance } from '../personnel-report/engine.mjs';
import { reportMessage, moscowDay, subscriptionAssignments } from './message.mjs';
import { telegramDirectorySnapshot, changeTelegramAccess } from './directory.mjs';
import { readHistoryMetrics } from '../personnel-report/history.mjs';
const BASE=Deno.env.get('SUPABASE_URL')!,KEY=Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
const cors={'Access-Control-Allow-Origin':'https://soltnar.github.io','Access-Control-Allow-Headers':'authorization,apikey,content-type','Access-Control-Allow-Methods':'GET,POST,OPTIONS','Cache-Control':'no-store'};
const json=(value:unknown,status=200)=>Response.json(value,{status,headers:cors});
async function botToken() {
  const environment=Deno.env.get('TELEGRAM_BOT_TOKEN');
  if(environment)return environment;
  const rows=await db('personnel_telegram_config?select=bot_token&id=eq.1');
  return rows[0]?.bot_token||'';
}
async function db(path:string,init:RequestInit={}) {
  const response=await fetch(`${BASE}/rest/v1/${path}`,{...init,signal:AbortSignal.timeout(15000),headers:{apikey:KEY,Authorization:`Bearer ${KEY}`,'Content-Type':'application/json',Prefer:'return=representation,resolution=merge-duplicates'}});
  if(!response.ok)throw new Error('Ошибка базы рассылки');
  return response.status===204?[]:await response.json();
}
async function summary(date:string, assignments:Array<{restaurant:string;roles:string[]}>=[]) {
  const days=await db(`attendance_daily_cache?select=work_date,status,rows&work_date=eq.${date}`);
  if(days.length!==1||days[0].status!=='ready')throw new Error('Проходная ещё не загружена полностью');
  const sales=await db(`revenue_entries?select=sale_date,restaurant,revenue&sale_date=eq.${date}&order=warehouse_id.asc&limit=1000`);
  if(!sales.length)throw new Error('Выручка ещё не загружена');
  if(assignments.some(item=>item.roles.some(role=>['hall','delivery','operators'].includes(role)))) {
    const history=await readHistoryMetrics(date,date,key=>Deno.env.get(key));
    return summarizeAttendance(days,sales,history.metrics,history.network,history.completeDates);
  }
  return summarizeAttendance(days,sales);
}
function dailyReader(date:string) {
  const cache=new Map<string,ReturnType<typeof summary>>();
  return (assignments:Array<{restaurant:string;roles:string[]}>)=>{
    const key=assignments.some(item=>item.roles.some(role=>['hall','delivery','operators'].includes(role)))?'history':'base';
    if(!cache.has(key))cache.set(key,summary(date,assignments));
    return cache.get(key)!;
  };
}
async function telegram(chat:string,text:string,messageId?:string) {
  const token=await botToken();
  if(!token)throw new Error('Добавьте TELEGRAM_BOT_TOKEN в секреты Supabase');
  // Never log the URL or the underlying fetch exception: they contain the token.
  let response;
  try {response=await fetch(`https://api.telegram.org/bot${token}/${messageId?'editMessageText':'sendMessage'}`,{method:'POST',signal:AbortSignal.timeout(12000),headers:{'Content-Type':'application/json'},body:JSON.stringify({chat_id:chat,text,...(messageId?{message_id:messageId}:{protect_content:true}),link_preview_options:{is_disabled:true}})});}
  catch {throw new Error('Статус доставки неизвестен; проверьте Telegram перед повтором');}
  const body=await response.json();
  if(messageId&&body.error_code===400&&String(body.description||'').includes('message is not modified'))return messageId;
  if(!response.ok||!body.ok)throw new Error(`Telegram отклонил отправку (код ${body.error_code||response.status})`);
  return String(body.result.message_id);
}
async function dispatch(date:string,recipientId?:string,scheduled=false,refresh=false,onlyRestaurants?:string[]) {
  if(!await botToken())return {sent:0,reason:'Бот ещё не подключён'};
  const {hour}=moscowDay();
  const recipients=await db(`personnel_telegram_recipients?select=*&enabled=eq.true${recipientId?'&id=eq.'+encodeURIComponent(recipientId):''}${!recipientId||scheduled?'&send_hour=lte.'+hour:''}&order=id.asc`);
  if(!recipients.length)return {sent:0,reason:'Нет активных получателей'};
  const result=[],readDaily=dailyReader(date);
  for(const recipient of recipients) {
    const known=await db(`personnel_telegram_log?select=id,status,message_id,chat_id,restaurant&work_date=eq.${date}&recipient_id=eq.${recipient.id}`);
    const assignments=subscriptionAssignments(recipient.assignments||[{restaurant:recipient.restaurant,roles:['kitchen']}]).filter(item=>!onlyRestaurants||onlyRestaurants.includes(item.restaurant)).sort((a,b)=>Number(known.some(log=>log.restaurant===a.restaurant))-Number(known.some(log=>log.restaurant===b.restaurant)));
    for(const [index,assignment] of assignments.entries()) {
      const restaurant=assignment.restaurant;
      const existing=known.filter(log=>log.restaurant===restaurant);
      if(existing.length&&!refresh)continue;
      if(result.filter(item=>!item.error).length>=3)return {result,hasMore:true,...(refresh?{remaining:assignments.slice(index).map(item=>item.restaurant)}:{})};
      let text;
      try{text=reportMessage(await readDaily([assignment]),restaurant,date,assignment.roles);}
      catch(error){result.push({id:recipient.id,restaurant,error:(error as Error).message});continue;}
      if(existing.length) {
        const log=existing[0];
        if(log.status!=='sent'||!log.message_id||log.chat_id!==recipient.chat_id){result.push({id:recipient.id,restaurant,status:'review',error:'Предыдущая доставка требует проверки. Автоматический повтор заблокирован.'});continue;}
        try {
          await telegram(recipient.chat_id,text,log.message_id);
          await db(`personnel_telegram_log?id=eq.${log.id}`,{method:'PATCH',body:JSON.stringify({roles:assignment.roles,error:null,updated_at:new Date().toISOString()})});
          result.push({id:recipient.id,restaurant,status:'updated'});
        }catch(error){result.push({id:recipient.id,restaurant,error:(error as Error).message});}
        continue;
      }
      // Each enterprise/day is claimed independently, so interrupted multi-venue
      // deliveries resume without repeating messages that already went out.
      const claim=await fetch(`${BASE}/rest/v1/personnel_telegram_log`,{method:'POST',headers:{apikey:KEY,Authorization:`Bearer ${KEY}`,'Content-Type':'application/json',Prefer:'return=representation'},body:JSON.stringify({recipient_id:recipient.id,restaurant,roles:assignment.roles,chat_id:recipient.chat_id,work_date:date,status:'sending'})});
      if(claim.status===409)continue;
      if(!claim.ok)throw new Error('Не удалось зарегистрировать отправку');
      const [log]=await claim.json();
      try {
        const messageId=await telegram(recipient.chat_id,text);
        await db(`personnel_telegram_log?id=eq.${log.id}`,{method:'PATCH',body:JSON.stringify({status:'sent',message_id:messageId,updated_at:new Date().toISOString()})});
        result.push({id:recipient.id,restaurant,status:'sent'});
      }catch(error){
        await db(`personnel_telegram_log?id=eq.${log.id}`,{method:'PATCH',body:JSON.stringify({status:'review',error:(error as Error).message,updated_at:new Date().toISOString()})});
        result.push({id:recipient.id,restaurant,status:'review'});
      }
    }
  }
  return {result};
}
Deno.serve(async(req)=>{
  if(req.method==='OPTIONS')return new Response(null,{headers:cors});
  try {
    const url=new URL(req.url),action=url.searchParams.get('action')||'settings';
    if(action==='cron') {
      if(req.method!=='GET')return json({error:'Метод не поддерживается'},405);
      const key=req.headers.get('x-cron-key');
      if(!key)return json({error:'Доступ запрещён'},403);
      const hash=[...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(key)))].map(v=>v.toString(16).padStart(2,'0')).join('');
      if(!(await db(`attendance_sync_config?select=id&id=eq.1&secret_hash=eq.${hash}`)).length)return json({error:'Доступ запрещён'},403);
      const recipientId=url.searchParams.get('recipient_id')||undefined;
      if(recipientId&&!/^[0-9a-f-]{36}$/.test(recipientId))return json({error:'Некорректный получатель'},400);
      return json(await dispatch(moscowDay().yesterday,recipientId,true));
    }
    const authorization=req.headers.get('authorization')||'';
    if(!authorization.startsWith('Bearer '))return json({error:'Войдите через Google'},401);
    const auth=await fetch(`${BASE}/auth/v1/user`,{headers:{authorization,apikey:KEY},signal:AbortSignal.timeout(10000)});
    if(!auth.ok)return json({error:'Сессия истекла'},401);
    const user=await auth.json();
    if(!user.email_confirmed_at||String(user.email).toLowerCase()!=='soltnar@gmail.com')return json({error:'Только владелец'},403);
    if(action==='directory'&&req.method==='GET')return json(await telegramDirectorySnapshot(authorization));
    if(action==='settings'&&req.method==='GET')return json({configured:!!await botToken(),recipients:await db('personnel_telegram_recipients?select=*&order=restaurant.asc'),log:await db('personnel_telegram_log?select=*&order=created_at.desc&limit=100')});
    if(req.method!=='POST')return json({error:'Метод не поддерживается'},405);
    const body=await req.json();
    if(action==='access')return json(await changeTelegramAccess(authorization,body));
    if(action==='save') {
      let assignments;
      try {assignments=subscriptionAssignments(body.assignments||[{restaurant:body.restaurant,roles:body.roles||['kitchen']}]);}
      catch(error){return json({error:(error as Error).message},400);}
      const chat=String(body.chat_id||'').trim(),name=String(body.name||'').trim(),hour=Number(body.send_hour??10);
      if(!/^[-]?\d{5,20}$/.test(chat)||!name||name.length>150||!Number.isInteger(hour)||hour<7||hour>18||typeof body.enabled!=='boolean')return json({error:'Проверьте имя, chat_id и час (7–18 МСК)'},400);
      if(body.id&&!/^[0-9a-f-]{36}$/.test(String(body.id)))return json({error:'Некорректный получатель'},400);
      const values={restaurant:assignments[0].restaurant,assignments,chat_id:chat,name,send_hour:hour,enabled:body.enabled,updated_at:new Date().toISOString()};
      return json({recipients:await db(body.id?`personnel_telegram_recipients?id=eq.${encodeURIComponent(body.id)}`:'personnel_telegram_recipients?on_conflict=chat_id',{method:body.id?'PATCH':'POST',body:JSON.stringify(values)})});
    }
    if(action==='preview'||action==='send') {
      const date=String(body.date||moscowDay().yesterday);
      if(!/^\d{4}-\d{2}-\d{2}$/.test(date)||date>=moscowDay().today)return json({error:'Выберите завершённый день'},400);
      if(action==='preview') {
        let assignments;
        try {assignments=subscriptionAssignments(body.assignments||[{restaurant:body.restaurant,roles:body.roles||['kitchen']}]);}
        catch(error){return json({error:(error as Error).message},400);}
        const messages=[],readDaily=dailyReader(date);
        for(const item of assignments){
          try {messages.push({restaurant:item.restaurant,text:reportMessage(await readDaily([item]),item.restaurant,date,item.roles)});}
          catch(error){messages.push({restaurant:item.restaurant,error:(error as Error).message});}
        }
        return json({messages,text:messages.map(item=>item.text||item.restaurant+' — '+item.error).join('\n\n────────\n\n')});
      }
      if(!/^[0-9a-f-]{36}$/.test(String(body.id||'')))return json({error:'Выберите получателя'},400);
      if(body.remaining!==undefined&&(!Array.isArray(body.remaining)||body.remaining.length>18||body.remaining.some((name:unknown)=>typeof name!=='string'||name.length>200)))return json({error:'Некорректный список предприятий'},400);
      return json(await dispatch(date,body.id,false,true,body.remaining));
    }
    return json({error:'Неизвестное действие'},400);
  }catch(error){return json({error:error instanceof Error?error.message:'Ошибка рассылки'},Number((error as {status?:number})?.status)||500);}
});
