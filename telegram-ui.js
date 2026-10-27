const notifyPanel=document.getElementById('telegramSettings');
const notifyLabels={kitchen:'Кухня',hall:'Зал',bar:'Бар',delivery:'Доставка',operators:'Операторы'};
const notifyStatus=document.getElementById('telegramStatus');
let notifyEpoch=0;
async function notifyRequest(action,body) {
  const session=await getReportSession();
  if(!session||!workspaceIsOwner())throw new Error('Войдите как владелец через Google');
  const response=await fetch(`${SUPABASE_URL}/functions/v1/personnel-notify?action=${action}`,{method:body?'POST':'GET',headers:{Authorization:`Bearer ${session.access_token}`,apikey:SUPABASE_PUBLISHABLE_KEY,'Content-Type':'application/json'},...(body?{body:JSON.stringify(body)}:{}),signal:AbortSignal.timeout(60000)});
  const result=await response.json();if(!response.ok)throw new Error(result.error||'Ошибка рассылки');return result;
}
function notifyError(error){if(workspaceIsOwner())notifyStatus.textContent=error.message;}
function notifyAssignments() {
  return [...document.querySelectorAll('#telegramScopes fieldset')].map(field=>({restaurant:field.dataset.venue,roles:[...field.querySelectorAll('input:checked')].map(input=>input.value)})).filter(item=>item.roles.length);
}
function setNotifyAssignments(assignments=[]) {
  for(const field of document.querySelectorAll('#telegramScopes fieldset'))for(const input of field.querySelectorAll('input'))input.checked=assignments.some(item=>item.restaurant===field.dataset.venue&&item.roles.includes(input.value));
}
function notifyScopeLabel(assignments) {return assignments.map(item=>`${item.restaurant}: ${item.roles.map(role=>notifyLabels[role]).join(', ')}`).join('; ');}
function resetNotifyForm(){document.getElementById('telegramForm').reset();document.getElementById('telegramRecipientId').value='';setNotifyAssignments();}
async function loadNotify() {
  const epoch=notifyEpoch,result=await notifyRequest('settings');
  if(!workspaceIsOwner()||epoch!==notifyEpoch)return;
  notifyStatus.textContent=result.configured?'Бот подключён. Для каждого пользователя выберите предприятия и разделы. Время — московское.':'Бот ещё не подключён.';
  const list=document.getElementById('telegramRecipients');list.replaceChildren();
  for(const recipient of result.recipients){
    const assignments=recipient.assignments||[{restaurant:recipient.restaurant,roles:['kitchen']}];
    const item=document.createElement('article'),label=document.createElement('p');label.textContent=`${recipient.name} (${recipient.chat_id}) · ${recipient.send_hour}:00 МСК · ${recipient.enabled?'включено':'выключено'}`;item.append(label);
    const scope=document.createElement('p');scope.textContent=notifyScopeLabel(assignments);item.append(scope);
    const edit=document.createElement('button');edit.type='button';edit.className='secondaryBtn';edit.textContent='Изменить';edit.onclick=()=>{document.getElementById('telegramRecipientId').value=recipient.id;setNotifyAssignments(assignments);document.getElementById('telegramName').value=recipient.name;document.getElementById('telegramChat').value=recipient.chat_id;document.getElementById('telegramHour').value=recipient.send_hour;document.getElementById('telegramEnabled').checked=recipient.enabled;document.getElementById('telegramName').focus();};item.append(edit);
    const send=document.createElement('button');send.type='button';send.className='secondaryBtn';send.textContent='Отправить за выбранный день';send.disabled=!recipient.enabled||!result.configured;
    send.onclick=async()=>{send.disabled=true;try{const outcome=await notifyRequest('send',{id:recipient.id,date:document.getElementById('telegramDate').value});await loadNotify();if(!workspaceIsOwner())return;const result=outcome.result||[],sent=result.filter(item=>item.status==='sent').length,errors=result.filter(item=>item.status==='review'||item.error).length;notifyStatus.textContent=outcome.reason||`Отправлено: ${sent}. Требуют проверки: ${errors}.${outcome.hasMore?' Остались предприятия: нажмите отправку ещё раз, уже отправленные сообщения не повторятся.':''}`;}catch(error){notifyError(error);}finally{send.disabled=false;}};item.append(send);list.append(item);
  }
  const log=document.getElementById('telegramLog');log.replaceChildren();
  for(const entry of result.log){const p=document.createElement('p');p.textContent=`${entry.work_date} · ${entry.restaurant} · ${(entry.roles||['kitchen']).map(role=>notifyLabels[role]).join(', ')} · ${entry.status==='sent'?'отправлено':entry.status==='sending'?'отправка начата — при задержке проверьте доставку':'нужна проверка'}${entry.error?' · '+entry.error:''}`;log.append(p);}
}
for(const venue of accessVenues.filter(name=>name!=='Фабрика кухни')) {
  const field=document.createElement('fieldset');field.dataset.venue=venue;
  const legend=document.createElement('legend');legend.textContent=venue;field.append(legend);
  const roles=venue==='Коминтерна 166 CALL CENTRE'?['kitchen','bar','delivery','operators']:['kitchen','hall','bar','delivery'];
  for(const role of roles){const label=document.createElement('label'),input=document.createElement('input');input.type='checkbox';input.value=role;const span=document.createElement('span');span.textContent=notifyLabels[role];label.append(input,span);field.append(label);}
  document.getElementById('telegramScopes').append(field);
}
notifyPanel.addEventListener('toggle',()=>{if(notifyPanel.open&&workspaceIsOwner())loadNotify().catch(notifyError);});
document.getElementById('telegramDate').value=yesterdayMoscow();
document.getElementById('telegramReload').onclick=()=>loadNotify().catch(notifyError);
document.getElementById('telegramForm').onsubmit=async(event)=>{event.preventDefault();const button=event.submitter;button.disabled=true;try{const assignments=notifyAssignments();if(!assignments.length)throw new Error('Выберите хотя бы одно предприятие и раздел');await notifyRequest('save',{id:document.getElementById('telegramRecipientId').value||undefined,assignments,name:document.getElementById('telegramName').value,chat_id:document.getElementById('telegramChat').value,send_hour:Number(document.getElementById('telegramHour').value),enabled:document.getElementById('telegramEnabled').checked});await loadNotify();}catch(error){notifyError(error);}finally{button.disabled=false;}};
document.getElementById('telegramPreview').onclick=async()=>{try{const epoch=notifyEpoch,result=await notifyRequest('preview',{assignments:notifyAssignments(),date:document.getElementById('telegramDate').value});if(workspaceIsOwner()&&epoch===notifyEpoch)document.getElementById('telegramMessage').textContent=result.text;}catch(error){notifyError(error);}};
new MutationObserver(()=>{if(!workspaceIsOwner()){notifyEpoch++;for(const id of ['telegramRecipients','telegramLog','telegramMessage'])document.getElementById(id).replaceChildren();resetNotifyForm();notifyStatus.textContent='';}}).observe(document.getElementById('usersNav'),{attributes:true,attributeFilter:['hidden']});
document.getElementById('telegramNew').onclick=resetNotifyForm;
