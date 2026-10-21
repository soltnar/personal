(async()=>{
  const status=document.getElementById('callbackStatus');
  const params=new URLSearchParams(location.search);
  history.replaceState(null,'',location.pathname);
  try{
    const saved=JSON.parse(sessionStorage.getItem('personnelSabyPending')||'null');
    sessionStorage.removeItem('personnelSabyPending');
    if(params.get('error'))throw new Error('Вход в Saby отменён.');
    if(!saved||saved.state!==params.get('state')||Date.now()-saved.createdAt>300000||!params.get('code'))throw new Error('Срок входа истёк. Начните вход заново.');
    const response=await fetch('https://wqxbnwcdkobgeyhdmqup.supabase.co/functions/v1/personnel-saby-auth?action=exchange',{method:'POST',cache:'no-store',headers:{'Content-Type':'application/json',apikey:'sb_publishable_WzfB8mJAOBXpeNWa34hBEQ_11QhCyqa'},body:JSON.stringify({state:saved.state,verifier:saved.verifier,code:params.get('code')})});
    const result=await response.json();if(!response.ok)throw new Error(result.error||'Saby не подтвердил вход');
    sessionStorage.setItem('personnelSabySession',JSON.stringify(result));
    location.replace('./?saby=login');
  }catch(error){status.textContent=error.message;}
})();
