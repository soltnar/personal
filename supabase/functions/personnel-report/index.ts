import { summarizeAttendance, restrictRows, roleKeys } from './engine.mjs';
const URL_BASE = Deno.env.get('SUPABASE_URL')!;
const KEY = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
const OWNER = 'soltnar@gmail.com';
const cors = {'Access-Control-Allow-Origin':'https://soltnar.github.io','Access-Control-Allow-Headers':'authorization,apikey,content-type','Access-Control-Allow-Methods':'GET,POST,OPTIONS','Cache-Control':'no-store'};
const json = (data: unknown,status=200)=>new Response(JSON.stringify(data),{status,headers:{...cors,'Content-Type':'application/json'}});
async function db(path:string, init:RequestInit={}) {
  const response=await fetch(`${URL_BASE}/rest/v1/${path}`,{...init,signal:AbortSignal.timeout(20000),headers:{apikey:KEY,Authorization:`Bearer ${KEY}`,'Content-Type':'application/json',Prefer:'return=representation,resolution=merge-duplicates'}});
  if(!response.ok) throw new Error('Ошибка чтения базы');
  return response.status===204?[]:await response.json();
}
async function pages(path:string) {
  const rows=[];
  for(let offset=0;offset<100000;offset+=1000){
    const page=await db(`${path}&limit=1000&offset=${offset}`);rows.push(...page);
    if(page.length<1000) return rows;
  }
  throw new Error('Превышен объём отчёта');
}
Deno.serve(async(req)=>{
  if(req.method==='OPTIONS') return new Response(null,{headers:cors});
  try {
    const authorization=req.headers.get('authorization')||'';
    if(!authorization.startsWith('Bearer ')) return json({error:'Войдите в приложение'},401);
    const auth=await fetch(`${URL_BASE}/auth/v1/user`,{headers:{authorization,apikey:KEY},signal:AbortSignal.timeout(15000)});
    if(!auth.ok) return json({error:'Сессия истекла'},401);
    const user=await auth.json();
    if(!user.email_confirmed_at) return json({error:'Подтвердите email'},403);
    const email=String(user.email||'').trim().toLowerCase(),admin=email===OWNER;
    const url=new URL(req.url),resource=url.searchParams.get('resource');
    if(resource==='access'){
      if(!admin) return json({error:'Только администратор'},403);
      if(req.method==='GET') return json({grants:await pages('personnel_access?select=email,restaurant,roles,updated_at&order=email.asc,restaurant.asc')});
      if(req.method!=='POST') return json({error:'Метод не поддерживается'},405);
      const body=await req.json();
      const target=String(body.email||'').trim().toLowerCase(),restaurant=String(body.restaurant||'').trim();
      if(!/^[a-z0-9._+@-]+$/.test(target)||!target.includes('@')||restaurant.length<3||restaurant.length>200||!Array.isArray(body.roles)||body.roles.some((r:string)=>!roleKeys.includes(r))) return json({error:'Проверьте email, предприятие и разделы'},400);
      if(body.roles.length===0) await db(`personnel_access?email=eq.${encodeURIComponent(target)}&restaurant=eq.${encodeURIComponent(restaurant)}`,{method:'DELETE'});
      else await db('personnel_access?on_conflict=email,restaurant',{method:'POST',body:JSON.stringify({email:target,restaurant,roles:[...new Set(body.roles)],updated_at:new Date().toISOString()})});
      return json({ok:true});
    }
    if(req.method!=='GET') return json({error:'Метод не поддерживается'},405);
    const grants=admin?[]:await db(`personnel_access?select=restaurant,roles&email=eq.${encodeURIComponent(email)}`);
    if(!admin&&!grants.length) return json({error:'Доступ не назначен. Обратитесь к администратору.'},403);
    const from=url.searchParams.get('from')||'',to=url.searchParams.get('to')||'';
    if(!/^\d{4}-\d{2}-\d{2}$/.test(from)||!/^\d{4}-\d{2}-\d{2}$/.test(to)||from>to||(Date.parse(to)-Date.parse(from))/86400000>61) return json({error:'Выберите период до 62 дней'},400);
    const [days,sales]=await Promise.all([
      pages(`attendance_daily_cache?select=work_date,status,rows,updated_at&work_date=gte.${from}&work_date=lte.${to}&order=work_date.asc`),
      pages(`revenue_entries?select=sale_date,restaurant,revenue&sale_date=gte.${from}&sale_date=lte.${to}&order=sale_date.asc,warehouse_id.asc`)
    ]);
    const missing=[];
    for(let d=from;d<=to;d=new Date(Date.parse(d)+86400000).toISOString().slice(0,10)) if(!days.some(day=>day.work_date===d&&day.status==='ready')) missing.push(d);
    if(missing.length) return json({error:`Не готовы дни проходной: ${missing.join(', ')}. Обновите период в настройках.`,missingDates:missing},409);
    const rows=restrictRows(summarizeAttendance(days,sales),grants,admin);
    return json({rows,admin,email,generatedAt:new Date().toISOString(),from,to});
  }catch(error){console.error(error);return json({error:'Не удалось получить отчёт. Повторите попытку.'},500);}
});
