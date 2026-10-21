import assert from 'node:assert/strict';
let handler,states=[],users=[],sessions=[],config={client_id:'interactive-test',client_secret:'secret-test'};
globalThis.Deno={env:{get:key=>key==='SUPABASE_URL'?'https://db.test':'server-key'},serve:callback=>{handler=callback;}};
globalThis.fetch=async(url,init={})=>{
  const u=new URL(url);let result;
  if(u.pathname==='/oauth/service/token'){
    assert.equal(init.headers['Content-Type'],'application/x-www-form-urlencoded');assert.equal(init.body.get('client_id'),'interactive-test');assert.equal(init.body.get('client_secret'),'secret-test');result={access_token:'saby-secret-token'};
  }else if(u.pathname==='/service/user_info'){assert.equal(init.headers['X-SBISAccessToken'],'saby-secret-token');result={'Пользователь':123,'ИдентификаторКлиента':456,'ИмяПользователя':'Пользователь Тест','ЛогинПользователя':'test',isService:false,guest:false,anonymous:false};}
  else if(u.pathname==='/auth/v1/user')result={email:'staff@example.org',email_confirmed_at:'2026-01-01'};
  else if(u.pathname.endsWith('personnel_saby_config'))result=config?[config]:[];
  else if(u.pathname.endsWith('personnel_saby_states')){
    if(init.method==='POST'){const value=JSON.parse(init.body);states.push(value);result=[value];}
    else if(init.method==='DELETE'){
      const matched=states.filter(s=>['state_hash','verifier_hash'].every(key=>!u.searchParams.has(key)||u.searchParams.get(key)==='eq.'+s[key])&&(!u.searchParams.has('expires_at')|| (u.searchParams.get('expires_at').startsWith('gt.')?s.expires_at>u.searchParams.get('expires_at').slice(3):s.expires_at<u.searchParams.get('expires_at').slice(3))));states=states.filter(s=>!matched.includes(s));result=matched;
    }else result=[];
  }else if(u.pathname.endsWith('personnel_saby_users')){
    if(init.method==='POST'){const value={enabled:false,...JSON.parse(init.body)};users.push(value);result=[value];}
    else result=users;
  }else if(u.pathname.endsWith('personnel_saby_sessions')){
    if(init.method==='POST'){const value=JSON.parse(init.body);sessions.push(value);result=[value];}else result=[];
  }else throw new Error('Unexpected path '+u.pathname);
  return new Response(JSON.stringify(result),{headers:{'Content-Type':'application/json'}});
};
await import('../supabase/functions/personnel-saby-auth/index.ts');
const request=(action,body)=>new Request('https://api.test/auth?action='+action,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
const {hash}=await import('../supabase/functions/personnel-saby-auth/identity.ts');
const verifier='a'.repeat(64);
let response=await handler(request('start',{verifierHash:await hash(verifier)}));assert.equal(response.status,200);const start=await response.json();
const url=new URL(start.url);assert.equal(url.origin,'https://online.sbis.ru');assert.equal(url.searchParams.get('response_type'),'code');assert(!start.url.includes('secret-test'));
response=await handler(request('exchange',{state:start.state,verifier:'b'.repeat(64),code:'code-test'}));assert.equal(response.status,400);assert.equal(users.length,0);
response=await handler(request('exchange',{state:start.state,verifier,code:'code-test'}));assert.equal(response.status,200);const session=await response.json();assert.equal(session.enabled,false);assert.equal(session.subject,'456:123');assert.equal(users[0].enabled,false);assert(!JSON.stringify(sessions).includes(session.token));assert(!JSON.stringify(sessions).includes('saby-secret-token'));
response=await handler(request('exchange',{state:start.state,verifier,code:'code-test'}));assert.equal(response.status,400);
response=await handler(request('admin',{operation:'enable',subject:'456:123',enabled:true}));assert.equal(response.status,403);
config=null;response=await handler(request('start',{verifierHash:await hash(verifier)}));assert.equal(response.status,409);
console.log('PASS: OAuth code flow, browser binding, one-time state, pending user without rights, hashed sessions, admin-only approval, unconfigured login disabled');
