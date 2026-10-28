const notifyPanel=document.getElementById('telegramSettings');
const notifyLabels={kitchen:'Кухня',hall:'Зал',bar:'Бар',delivery:'Доставка',operators:'Операторы'};
const notifyStatus=document.getElementById('telegramStatus');
let notifyEpoch=0,notifyDirectory=[],notifyRecipients=[];
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
function renderNotifyDirectory(selected='') {
  const picker=document.getElementById('telegramUser');picker.replaceChildren();
  const empty=document.createElement('option');empty.value='';empty.textContent='Выберите пользователя Telegram';picker.append(empty);
  for(const user of notifyDirectory){const option=document.createElement('option');option.value=user.id;option.textContent=user.name+' · '+user.id+(user.status==='approved'?'':' · доступ не разрешён');option.disabled=user.status!=='approved';picker.append(option);}
  if(selected&&!notifyDirectory.some(user=>user.id===selected)){const option=document.createElement('option');option.value=selected;option.textContent='Сохранённый получатель · '+selected;picker.append(option);}
  picker.value=selected;
}
function showNotifySource(id) {
  const user=notifyDirectory.find(user=>user.id===id);
  document.getElementById('telegramSourceRights').textContent=user?'Права в выручке: '+user.accessLabel+'. Кухню и бар можно подставить; зал, доставка и операторы назначаются отдельно.':'Выберите пользователя из приложения выручки. Сохранённые настройки рассылки не меняются при обновлении списка.';
  document.getElementById('telegramImportRights').disabled=!user||user.status!=='approved'||!user.suggestedAssignments.length;
}
function resetNotifyForm(){document.getElementById('telegramForm').reset();document.getElementById('telegramRecipientId').value='';setNotifyAssignments();document.getElementById('telegramUser').value='';showNotifySource('');}
async function loadNotify() {
  const epoch=notifyEpoch;const [settings,directory]=await Promise.allSettled([notifyRequest('settings'),notifyRequest('directory')]);
  if(settings.status==='rejected')throw settings.reason;
  const result=settings.value;
  if(!workspaceIsOwner()||epoch!==notifyEpoch)return;
  const selected=document.getElementById('telegramChat').value;
  if(directory.status==='fulfilled')notifyDirectory=directory.value.users;
  renderNotifyDirectory(selected);showNotifySource(selected);
  document.getElementById('telegramDirectoryStatus').textContent=directory.status==='fulfilled'?'Пользователи и права загружены из приложения выручки.':directory.reason.message;
  notifyStatus.textContent=result.configured?'Бот подключён. Для каждого пользователя выберите предприятия и разделы. Время — московское.':'Бот ещё не подключён.';
  notifyRecipients=result.recipients;
  const list=document.getElementById('telegramRecipients');list.replaceChildren();
  for(const recipient of result.recipients){
    const assignments=recipient.assignments||[{restaurant:recipient.restaurant,roles:['kitchen']}];
    const item=document.createElement('article');item.className='telegramRecipientRow';
    const identity=document.createElement('div');identity.className='telegramRecipientIdentity';const label=document.createElement('strong');label.textContent=recipient.name;const id=document.createElement('span');id.textContent='Telegram '+recipient.chat_id;identity.append(label,id);item.append(identity);
    const scope=document.createElement('div');scope.className='telegramRecipientScope';scope.textContent='→ '+notifyScopeLabel(assignments);item.append(scope);
    const state=document.createElement('span');state.className='telegramRecipientState';state.textContent=recipient.send_hour+':00 МСК · '+(recipient.enabled?'включено':'выключено');item.append(state);
    const actions=document.createElement('div');actions.className='telegramRecipientActions';
    const edit=document.createElement('button');edit.type='button';edit.className='secondaryBtn';edit.textContent='Изменить';edit.onclick=()=>{document.getElementById('telegramRecipientId').value=recipient.id;setNotifyAssignments(assignments);document.getElementById('telegramName').value=recipient.name;document.getElementById('telegramChat').value=recipient.chat_id;renderNotifyDirectory(recipient.chat_id);showNotifySource(recipient.chat_id);document.getElementById('telegramHour').value=recipient.send_hour;document.getElementById('telegramEnabled').checked=recipient.enabled;document.getElementById('telegramUser').focus();};actions.append(edit);
    const send=document.createElement('button');send.type='button';send.className='secondaryBtn';send.textContent='Отправить / обновить сводку';send.disabled=!recipient.enabled||!result.configured;
    send.onclick=async()=>{send.disabled=true;const epoch=notifyEpoch;document.getElementById('telegramMessage').replaceChildren();try{const date=document.getElementById('telegramDate').value,outcome={result:[],reason:''};let remaining;do{const part=await notifyRequest('send',{id:recipient.id,date,remaining});outcome.result.push(...(part.result||[]));outcome.reason=part.reason||'';remaining=part.hasMore?part.remaining:undefined;if(!workspaceIsOwner()||epoch!==notifyEpoch)return;}while(remaining?.length);await loadNotify();if(!workspaceIsOwner())return;const result=outcome.result||[],sent=result.filter(item=>item.status==='sent').length,updated=result.filter(item=>item.status==='updated').length,errors=result.filter(item=>item.status==='review'||item.error).length;notifyStatus.textContent=outcome.reason||`Отправлено: ${sent}. Обновлено: ${updated}. Требуют проверки: ${errors}.`;if(sent||updated){const preview=await notifyRequest('preview',{assignments,date:document.getElementById('telegramDate').value});if(workspaceIsOwner()&&epoch===notifyEpoch)document.getElementById('telegramMessage').textContent=recipient.name+'\n\n'+preview.text;}}catch(error){notifyError(error);}finally{send.disabled=false;}};actions.append(send);item.append(actions);list.append(item);
  }
  const log=document.getElementById('telegramLog');log.replaceChildren();
  for(const entry of result.log){const p=document.createElement('p');p.textContent=`${result.recipients.find(recipient=>recipient.id===entry.recipient_id)?.name||entry.chat_id} → ${entry.work_date} · ${entry.restaurant} · ${(entry.roles||['kitchen']).map(role=>notifyLabels[role]).join(', ')} · ${entry.status==='sent'?'отправлено':entry.status==='sending'?'отправка начата — при задержке проверьте доставку':'нужна проверка'}${entry.error?' · '+entry.error:''}`;log.append(p);}
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
new MutationObserver(()=>{if(!workspaceIsOwner()){notifyEpoch++;notifyDirectory=[];notifyRecipients=[];renderNotifyDirectory();document.getElementById('telegramDirectoryStatus').textContent='';for(const id of ['telegramRecipients','telegramLog','telegramMessage'])document.getElementById(id).replaceChildren();resetNotifyForm();notifyStatus.textContent='';}}).observe(document.getElementById('usersNav'),{attributes:true,attributeFilter:['hidden']});
document.getElementById('telegramNew').onclick=resetNotifyForm;

document.getElementById('telegramUser').onchange=()=>{
  const id=document.getElementById('telegramUser').value,user=notifyDirectory.find(item=>item.id===id);
  if(!user||user.status!=='approved')return;
  resetNotifyForm();document.getElementById('telegramUser').value=id;
  document.getElementById('telegramName').value=user.name;document.getElementById('telegramChat').value=id;
  const existing=notifyRecipients.find(item=>item.chat_id===id);
  if(existing){document.getElementById('telegramRecipientId').value=existing.id;document.getElementById('telegramHour').value=existing.send_hour;document.getElementById('telegramEnabled').checked=existing.enabled;setNotifyAssignments(existing.assignments||[{restaurant:existing.restaurant,roles:['kitchen']}]);}
  else setNotifyAssignments(user.suggestedAssignments.filter(item=>accessVenues.includes(item.restaurant)));
  showNotifySource(id);
};
document.getElementById('telegramImportRights').onclick=()=>{
  const user=notifyDirectory.find(item=>item.id===document.getElementById('telegramUser').value);
  if(!user||user.status!=='approved')return;
  setNotifyAssignments(user.suggestedAssignments.filter(item=>accessVenues.includes(item.restaurant)));
  document.getElementById('telegramMessage').replaceChildren();
  notifyStatus.textContent='Подставлены кухня и бар из прав выручки. Проверьте разделы и сохраните получателя. Зал и операторы выбираются отдельно.';
};
