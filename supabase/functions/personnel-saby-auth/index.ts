import { hash } from './identity.ts';
const BASE=Deno.env.get('SUPABASE_URL')!,KEY=Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
const OWNER='soltnar@gmail.com',CALLBACK='https://soltnar.github.io/personal/saby-callback.html';
const cors={'Access-Control-Allow-Origin':'https://soltnar.github.io','Access-Control-Allow-Headers':'authorization,apikey,content-type','Access-Control-Allow-Methods':'GET,POST,OPTIONS','Cache-Control':'no-store'};
const json=(data:unknown,status=200)=>new Response(JSON.stringify(data),{status,headers:{...cors,'Content-Type':'application/json'}});
async function db(path:string,init:RequestInit={}){
  const response=await fetch(`${BASE}/rest/v1/${path}`,{...init,signal:AbortSignal.timeout(15000),headers:{apikey:KEY,Authorization:`Bearer ${KEY}`,'Content-Type':'application/json',Prefer:'return=representation,resolution=merge-duplicates'}});
  if(!response.ok)throw new Error('Ошибка базы авторизации');
  return response.status===204?[]:await response.json();
}
async function owner(req:Request){const auth=await fetch(`${BASE}/auth/v1/user`,{headers:{authorization:req.headers.get('authorization')||'',apikey:KEY}});if(!auth.ok)return false;const user=await auth.json();return user.email_confirmed_at&&String(user.email).toLowerCase()===OWNER;}
const random=()=>[...crypto.getRandomValues(new Uint8Array(32))].map(b=>b.toString(16).padStart(2,'0')).join('');
Deno.serve(async(req)=>{
  if(req.method==='OPTIONS')return new Response(null,{headers:cors});
  try{
    const action=new URL(req.url).searchParams.get('action')||'';
    if(req.method==='GET'&&action==='status'){const configs=await db('personnel_saby_config?id=eq.1&select=client_id');return json({configured:configs.length===1,callback:CALLBACK});}
    if(action==='admin'){
      if(!await owner(req))return json({error:'Только администратор Google'},403);
      if(req.method==='GET')return json({users:await db('personnel_saby_users?select=subject,display_name,login,enabled,last_login_at&order=last_login_at.desc&limit=1000'),grants:await db('personnel_saby_access?select=subject,restaurant,roles&limit=1000'),config:await db('personnel_saby_config?id=eq.1&select=client_id'),callback:CALLBACK});
      if(req.method!=='POST')return json({error:'Метод не поддерживается'},405);
      const body=await req.json();
      if(body.operation==='config'){
        if(typeof body.clientId!=='string'||typeof body.clientSecret!=='string'||!body.clientId.trim()||!body.clientSecret.trim()||body.clientId.length>500||body.clientSecret.length>2000)return json({error:'Укажите client_id и client_secret интерактивного приложения Saby'},400);
        await db('personnel_saby_config?on_conflict=id',{method:'POST',body:JSON.stringify({id:1,client_id:body.clientId.trim(),client_secret:body.clientSecret.trim(),updated_at:new Date().toISOString()})});return json({ok:true});
      }
      const subject=String(body.subject||'');
      if(!/^\d+:\d+$/.test(subject))return json({error:'Некорректный пользователь Saby'},400);
      if(body.operation==='enable'){
        if(typeof body.enabled!=='boolean')return json({error:'Некорректный статус'},400);
        await db(`personnel_saby_users?subject=eq.${encodeURIComponent(subject)}`,{method:'PATCH',body:JSON.stringify({enabled:body.enabled})});
        if(!body.enabled)await db(`personnel_saby_sessions?subject=eq.${encodeURIComponent(subject)}`,{method:'DELETE'});
        return json({ok:true});
      }
      if(body.operation==='grant'){
        const roles=body.roles,restaurant=String(body.restaurant||'').trim();
        if(!Array.isArray(roles)||roles.some(r=>!['kitchen','hall','bar','delivery'].includes(r))||restaurant.length<3||restaurant.length>200)return json({error:'Проверьте предприятие и разделы'},400);
        if(!roles.length)await db(`personnel_saby_access?subject=eq.${encodeURIComponent(subject)}&restaurant=eq.${encodeURIComponent(restaurant)}`,{method:'DELETE'});
        else await db('personnel_saby_access?on_conflict=subject,restaurant',{method:'POST',body:JSON.stringify({subject,restaurant,roles:[...new Set(roles)]})});
        return json({ok:true});
      }
      return json({error:'Неизвестная операция'},400);
    }
    if(req.method!=='POST')return json({error:'Метод не поддерживается'},405);
    const body=await req.json();
    if(action==='logout'){const token=(req.headers.get('authorization')||'').replace(/^Bearer /,'');if(token.startsWith('ps_'))await db(`personnel_saby_sessions?token_hash=eq.${await hash(token)}`,{method:'DELETE'});return json({ok:true});}
    if(action==='start'){
      if(!/^[a-f0-9]{64}$/.test(body.verifierHash||''))return json({error:'Некорректный запрос входа'},400);
      const config=(await db('personnel_saby_config?id=eq.1&select=client_id'))[0];if(!config)return json({error:'Администратор ещё не настроил вход Saby'},409);
      const now=new Date().toISOString();
      const ipHash=await hash(req.headers.get('x-real-ip')||(req.headers.get('x-forwarded-for')||'unknown').split(',')[0].trim());
      const recent=await db(`personnel_saby_states?ip_hash=eq.${ipHash}&expires_at=gt.${encodeURIComponent(now)}&select=state_hash&limit=20`);
      if(recent.length>=20)return json({error:'Слишком много попыток входа. Повторите через пять минут.'},429);
      await db(`personnel_saby_states?expires_at=lt.${encodeURIComponent(now)}`,{method:'DELETE'});
      await db(`personnel_saby_sessions?expires_at=lt.${encodeURIComponent(now)}`,{method:'DELETE'});
      const state=random();await db('personnel_saby_states',{method:'POST',body:JSON.stringify({state_hash:await hash(state),verifier_hash:body.verifierHash,ip_hash:ipHash,expires_at:new Date(Date.now()+300000).toISOString()})});
      const url=new URL('https://online.sbis.ru/oauth/api/token');url.search=new URLSearchParams({client_id:config.client_id,redirect_uri:CALLBACK,response_type:'code',scope:'openid profile',state}).toString();return json({url:url.href,state});
    }
    if(action!=='exchange')return json({error:'Неизвестная операция'},400);
    if(!/^[a-f0-9]{64}$/.test(body.state||'')||!/^[a-f0-9]{64}$/.test(body.verifier||'')||typeof body.code!=='string'||!body.code||body.code.length>4096)return json({error:'Некорректный ответ Saby'},400);
    // Atomic consumption: expired/reused states and another browser's verifier fail closed.
    const states=await db(`personnel_saby_states?state_hash=eq.${await hash(body.state)}&verifier_hash=eq.${await hash(body.verifier)}&expires_at=gt.${encodeURIComponent(new Date().toISOString())}`,{method:'DELETE'});
    if(states.length!==1)return json({error:'Вход истёк или уже использован. Начните заново.'},400);
    const config=(await db('personnel_saby_config?id=eq.1&select=client_id,client_secret'))[0];
    if(!config)return json({error:'Вход Saby не настроен'},409);
    const exchange=await fetch('https://online.sbis.ru/oauth/service/token',{method:'POST',signal:AbortSignal.timeout(20000),headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({code:body.code,client_id:config.client_id,client_secret:config.client_secret})});
    const tokens=await exchange.json();const token=tokens.access_token||tokens.token;
    if(!exchange.ok||typeof token!=='string'||!token)return json({error:'Saby не подтвердил вход. Проверьте настройки приложения и адрес возврата.'},401);
    const profile=await fetch('https://online.sbis.ru/service/user_info',{headers:{'X-SBISAccessToken':token},signal:AbortSignal.timeout(15000)});
    const info=await profile.json();
    if(!profile.ok||info.isService||info.anonymous||info.guest||!/^\d+$/.test(String(info['Пользователь']))||!/^\d+$/.test(String(info['ИдентификаторКлиента'])))return json({error:'Не удалось подтвердить личный аккаунт Saby'},401);
    const subject=`${info['ИдентификаторКлиента']}:${info['Пользователь']}`;
    const record={subject,display_name:String(info['ИмяПользователя']||[info['Фамилия'],info['Имя']].filter(Boolean).join(' ')||subject),login:String(info['ЛогинПользователя']||''),last_login_at:new Date().toISOString()};
    const existing=(await db(`personnel_saby_users?subject=eq.${encodeURIComponent(subject)}&select=enabled`))[0];
    if(existing)await db(`personnel_saby_users?subject=eq.${encodeURIComponent(subject)}`,{method:'PATCH',body:JSON.stringify(record)});
    else await db('personnel_saby_users',{method:'POST',body:JSON.stringify(record)});
    const session='ps_'+random(),expiresAt=new Date(Date.now()+8*3600000).toISOString();
    await db('personnel_saby_sessions',{method:'POST',body:JSON.stringify({token_hash:await hash(session),subject,expires_at:expiresAt})});
    return json({token:session,subject,name:record.display_name,enabled:existing?.enabled===true,expiresAt});
  }catch{ return json({error:'Не удалось выполнить вход. Повторите попытку или проверьте настройки Saby.'},500); }
});
