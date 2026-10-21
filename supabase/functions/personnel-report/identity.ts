const BASE=Deno.env.get('SUPABASE_URL')!;
const KEY=Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!;
export async function hash(value:string){return [...new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(value)))].map(b=>b.toString(16).padStart(2,'0')).join('');}
export async function readSabyIdentity(authorization:string){
  if(!authorization.startsWith('Bearer ps_'))return null;
  const response=await fetch(`${BASE}/rest/v1/personnel_saby_sessions?select=subject&token_hash=eq.${await hash(authorization.slice(7))}&expires_at=gt.${encodeURIComponent(new Date().toISOString())}`,{headers:{apikey:KEY,Authorization:`Bearer ${KEY}`}});
  if(!response.ok)return null;
  const sessions=await response.json();if(sessions.length!==1)return null;
  const userResponse=await fetch(`${BASE}/rest/v1/personnel_saby_users?subject=eq.${encodeURIComponent(sessions[0].subject)}&enabled=eq.true&select=subject,display_name`,{headers:{apikey:KEY,Authorization:`Bearer ${KEY}`}});
  if(!userResponse.ok)return null;const users=await userResponse.json();return users.length===1?users[0]:null;
}
