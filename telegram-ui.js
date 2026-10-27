const notifyPanel=document.getElementById('telegramSettings');
async function notifyRequest(action,body) {
  const session=await getReportSession();
  if(!session||!workspaceIsOwner())throw new Error('Войдите как владелец через Google');
  const response=await fetch(`${SUPABASE_URL}/functions/v1/personnel-notify?action=${action}`,{method:body?'POST':'GET',headers:{Authorization:`Bearer ${session.access_token}`,apikey:SUPABASE_PUBLISHABLE_KEY,'Content-Type':'application/json'},...(body?{body:JSON.stringify(body)}:{}),signal:AbortSignal.timeout(60000)});
  const result=await response.json();if(!response.ok)throw new Error(result.error||'Ошибка рассылки');return result;
}
const notifyStatus=document.getElementById('telegramStatus');
function notifyError(error){notifyStatus.textContent=error.message;}
async function loadNotify() {
  const result=await notifyRequest('settings');
  if(!workspaceIsOwner())return;
  notifyStatus.textContent=result.configured?'Бот подключён. Время отправки — московское.':'Бот ещё не подключён: добавьте TELEGRAM_BOT_TOKEN в секреты Supabase.';
  const list=document.getElementById('telegramRecipients');list.replaceChildren();
  for(const recipient of result.recipients){
    const item=document.createElement('p'),label=document.createElement('span');label.textContent=`${recipient.restaurant} → ${recipient.name} (${recipient.chat_id}), ${recipient.send_hour}:00 МСК · ${recipient.enabled?'включено':'выключено'} `;item.append(label);
    const edit=document.createElement('button');edit.type='button';edit.className='secondaryBtn';edit.textContent='Изменить';edit.onclick=()=>{document.getElementById('telegramRecipientId').value=recipient.id;document.getElementById('telegramVenue').value=recipient.restaurant;document.getElementById('telegramName').value=recipient.name;document.getElementById('telegramChat').value=recipient.chat_id;document.getElementById('telegramHour').value=recipient.send_hour;document.getElementById('telegramEnabled').checked=recipient.enabled;};item.append(edit);
    const send=document.createElement('button');send.type='button';send.className='secondaryBtn';send.textContent='Отправить за выбранный день';send.disabled=!recipient.enabled||!result.configured;
    send.onclick=async()=>{send.disabled=true;try{const outcome=await notifyRequest('send',{id:recipient.id,date:document.getElementById('telegramDate').value});await loadNotify();notifyStatus.textContent=outcome.reason||((outcome.result||[]).some(item=>item.status==='sent')?'Сообщение отправлено.':(outcome.result||[]).some(item=>item.status==='review')?'Проверьте доставку в Telegram и запись в журнале.':'За этот день уже есть запись в журнале; повтор не выполнен.');}catch(error){notifyError(error);}finally{send.disabled=false;}};item.append(send);list.append(item);
  }
  const log=document.getElementById('telegramLog');log.replaceChildren();
  for(const entry of result.log){const p=document.createElement('p');p.textContent=`${entry.work_date} · ${entry.restaurant} · ${entry.status==='sent'?'отправлено':entry.status==='sending'?'отправка начата — при задержке проверьте доставку':'нужна проверка'}${entry.error?' · '+entry.error:''}`;log.append(p);}
}
notifyPanel.addEventListener('toggle',()=>{if(notifyPanel.open&&workspaceIsOwner())loadNotify().catch(notifyError);});
for(const name of accessVenues.filter(name=>name!=='Фабрика кухни')){const option=document.createElement('option');option.value=name;option.textContent=name;document.getElementById('telegramVenue').append(option);}
document.getElementById('telegramDate').value=yesterdayMoscow();
document.getElementById('telegramReload').onclick=()=>loadNotify().catch(notifyError);
document.getElementById('telegramForm').onsubmit=async(event)=>{event.preventDefault();const button=event.submitter;button.disabled=true;try{await notifyRequest('save',{id:document.getElementById('telegramRecipientId').value||undefined,restaurant:document.getElementById('telegramVenue').value,name:document.getElementById('telegramName').value,chat_id:document.getElementById('telegramChat').value,send_hour:Number(document.getElementById('telegramHour').value),enabled:document.getElementById('telegramEnabled').checked});await loadNotify();}catch(error){notifyError(error);}finally{button.disabled=false;}};
document.getElementById('telegramPreview').onclick=async()=>{try{const result=await notifyRequest('preview',{restaurant:document.getElementById('telegramVenue').value,date:document.getElementById('telegramDate').value});if(workspaceIsOwner())document.getElementById('telegramMessage').textContent=result.text;}catch(error){notifyError(error);}};
// Financial data and recipient identities are cleared as soon as the owner signs out.
new MutationObserver(()=>{if(!workspaceIsOwner()){for(const id of ['telegramRecipients','telegramLog','telegramMessage'])document.getElementById(id).replaceChildren();document.getElementById('telegramForm').reset();notifyStatus.textContent='';}}).observe(document.getElementById('usersNav'),{attributes:true,attributeFilter:['hidden']});

document.getElementById('telegramNew').onclick=()=>{document.getElementById('telegramForm').reset();document.getElementById('telegramRecipientId').value='';};
