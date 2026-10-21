const REPORT_API = `${SUPABASE_URL}/functions/v1/personnel-report`;
const roleLabels = { kitchen:"Кухня", hall:"Зал", bar:"Бар", delivery:"Доставка" };
const roleRevenue = {kitchen:"kitchenRevenue",bar:"barRevenue",hall:"revenue",delivery:"revenue"};
async function reportRequest(path="", init={}) {
  const session=await getReportSession();
  if(!session) throw new Error("Войдите через Google или Saby");
  const response=await fetch(REPORT_API+path,{...init,cache:"no-store",signal:AbortSignal.timeout(60000),headers:{Authorization:`Bearer ${session.access_token}`,apikey:SUPABASE_PUBLISHABLE_KEY,"Content-Type":"application/json"}});
  const body=await response.json();
  if(!response.ok) throw new Error(body.error||"Не удалось получить данные");
  return body;
}
async function loadScopedReport() {
  if(!supabaseClient) return;
  const startingSession=await getReportSession();
  const startingUser=startingSession?.user?.id;
  const from=sabyFromInput.value,to=sabyToInput.value;
  if(!from||!to||from>to) return setSabyApiStatus("Проверьте период","error");
  loadSabyBtn.disabled=true;loadSabyBtn.classList.add("is-busy");
  try {
    setSabyApiStatus("Читаем разрешённые данные из базы…","loading");
    const payload=await reportRequest(`?from=${from}&to=${to}`);
    const currentUser=(await getReportSession())?.user?.id;
    if(!startingUser||currentUser!==startingUser) return;
    reportRows=payload.rows;reportAdmin=payload.admin;
    document.getElementById("accessSettings").hidden=!reportAdmin;
    document.getElementById("forceSyncBtn").hidden=!reportAdmin;
    const restaurants=[...new Set(reportRows.map(r=>r.restaurant))].sort((a,b)=>a.localeCompare(b,"ru"));
    fillMultiSelect(restaurantSelect,restaurants);
    fillMultiSelect(dateSelect,[...new Set(reportRows.map(r=>r.dateIso))].sort());
    document.querySelectorAll('.groupCheck').forEach(input=>{
      const role=Object.keys(roleLabels).find(k=>roleLabels[k]===input.value);
      input.disabled=!reportAdmin&&!reportRows.some(r=>r.allowedRoles.includes(role));
      input.checked=!input.disabled;input.closest('label').hidden=input.disabled;
    });
    const select=document.getElementById("accessRestaurant");select.replaceChildren();
    restaurants.filter(r=>r!=="Подразделение не определено").forEach(r=>{const o=document.createElement('option');o.value=r;o.textContent=r;select.append(o);});
    mappedRecords=[];baseRecords=[];
    lastResultRows=calculateScopedRows();renderTable(lastResultRows);
    setSabyApiStatus(`Отчёт из базы · ${from} — ${to} · ${payload.email}`,"success");
    setRevenueDbStatus("Выработка: ₽ на учитываемую смену. Общий показатель = сумма выручки / сумма смен.","success");
    statusEl.textContent="Доступ проверен сервером. Отображаются только разрешённые предприятия и разделы.";
  }catch(error){
    reportRows=[];lastResultRows=[];renderTable([]);
    csvBtn.disabled=true;xlsxBtn.disabled=true;
    setSabyApiStatus(error.message,"error");
  }finally{loadSabyBtn.disabled=false;loadSabyBtn.classList.remove("is-busy");}
}
function calculateScopedRows() {
  const selected=new Set(getSelectedValues(restaurantSelect)),groups=new Set(getCheckedGroups());
  const daily=(reportRows||[]).filter(r=>selected.has(r.restaurant)).map(source=>{
    const row={...source,details:{...source.details},total:0,hasConflict:false};
    for(const [role,label] of Object.entries(roleLabels)) {
      if(!groups.has(label)||!source.allowedRoles.includes(role)){row[role]=null;row.details[role]=[];}
      else row.total+=row[role]||0;
      if(groups.has(label)&&source.allowedRoles.includes(role)&&row.details[role]?.some(p=>p.hasConflict)) row.hasConflict=true;
    }
    row.coverage={};
    for(const role of Object.keys(roleLabels)) row.coverage[role]=productivitySample(row,role);
    return row;
  });
  if(document.getElementById("reportView").value==="daily") return daily;
  const map=new Map();
  for(const row of daily){
    if(!map.has(row.restaurant)) map.set(row.restaurant,{...row,dateIso:"",dailyRows:[],details:{kitchen:[],hall:[],bar:[],delivery:[]},coverage:Object.fromEntries(Object.keys(roleLabels).map(k=>[k,{revenue:0,shifts:0,covered:0,missing:0,missingShifts:0}])),kitchen:row.kitchen===null?null:0,hall:row.hall===null?null:0,bar:row.bar===null?null:0,delivery:row.delivery===null?null:0,revenue:null,kitchenRevenue:null,barRevenue:null,total:0});
    const target=map.get(row.restaurant);target.total+=row.total;target.hasConflict ||= row.hasConflict;
    if(row.kitchenRevenueSource==='unified')target.kitchenRevenueSource='unified';
    target.dailyRows.push(row);
    for(const role of Object.keys(roleLabels)){if(Number.isFinite(row[role]))target[role]+=row[role];target.details[role].push(...(row.details[role]||[]).map(p=>({...p,dateIso:row.dateIso})));}
    for(const field of ['revenue','kitchenRevenue','barRevenue']) if(Number.isFinite(row[field])) target[field]=(target[field]??0)+row[field];
    for(const role of Object.keys(roleLabels)) for(const key of ['revenue','shifts','covered','missing','missingShifts']) target.coverage[role][key]+=row.coverage[role][key];
  }
  return [...map.values()];
}
function formatProductivity(row,role) {
  const sample=row.coverage?.[role]||productivitySample(row,role);
  return sample.shifts>0?`${formatMoney(sample.revenue/sample.shifts)} ₽${sample.missing?' *':''}`:"—";
}
function productivitySample(row,role){
  const shifts=row[role]||0,revenue=row[roleRevenue[role]];
  const valid=Number.isFinite(revenue)&&shifts>0;
  return {revenue:valid?revenue:0,shifts:valid?shifts:0,covered:valid?1:0,missing:shifts>0&&!Number.isFinite(revenue)?1:0,missingShifts:shifts>0&&!Number.isFinite(revenue)?shifts:0};
}
function roleMetricsHtml(row){
  return `<div class="dayMetrics">${Object.entries(roleLabels).filter(([role,label])=>getCheckedGroups().includes(label)&&row.allowedRoles?.includes(role)).map(([role,label])=>{
    const sample=row.coverage?.[role]||productivitySample(row,role),revenue=row[roleRevenue[role]];
    return `<article><h4>${label}</h4><dl><div><dt>${role==='kitchen'?(row.kitchenRevenueSource==='unified'?'Выручка единого склада':'Выручка кухни'):role==='bar'?'Выручка бара':'Общая выручка'}</dt><dd>${Number.isFinite(revenue)?formatMoney(revenue)+' ₽':'нет данных'}</dd></div><div><dt>Смены</dt><dd>${formatShift(row[role])}</dd></div><div><dt>Выработка</dt><dd>${formatProductivity(row,role)}</dd></div></dl>${sample.missing?`<p class="coverageWarning">Без выручки: ${sample.missing} предприятие-дней / ${formatShift(sample.missingShifts)} смен. В расчёте: ${formatShift(sample.shifts)} смен.</p>`:''}</article>`;
  }).join('')}</div>`;
}
function renderGaps(rows){
  const container=document.getElementById('reportGaps');if(!container)return;
  container.replaceChildren();
  const daily=rows.flatMap(r=>r.dailyRows||[r]);
  const gaps=[];
  for(const row of daily)for(const [role,label] of Object.entries(roleLabels)){
    if(!row.allowedRoles?.includes(role)||!getCheckedGroups().includes(label))continue;
    const sample=row.coverage?.[role]||productivitySample(row,role);
    if(sample.missing)gaps.push({row,role,label,shifts:sample.missingShifts});
  }
  const heading=document.createElement('summary');heading.textContent=`Пропуски выручки — ${gaps.length} записей. Показать предприятия и даты`;container.append(heading);
  if(!gaps.length){const p=document.createElement('p');p.textContent='Пропусков выручки в выбранных разделах со сменами нет.';container.append(p);return;}
  const intro=document.createElement('p');intro.className='hint';intro.textContent='Эти дни не вошли в выработку соответствующей группы. «Открыть день» покажет смены и сотрудников.';container.append(intro);
  const list=document.createElement('div');list.className='gapList';
  gaps.sort((a,b)=>a.row.restaurant.localeCompare(b.row.restaurant,'ru')||a.row.dateIso.localeCompare(b.row.dateIso));
  for(const gap of gaps){
    const item=document.createElement('article');
    const content=document.createElement('div');
    const name=document.createElement('strong');name.textContent=`${gap.row.restaurant} · ${prettyDate(gap.row.dateIso)} · ${gap.label}`;
    const reason=document.createElement('p');reason.textContent=`${formatShift(gap.shifts)} смен · ${gap.row.restaurant==='Фабрика кухни'?'У производственного предприятия нет собственной выручки':gap.role==='kitchen'?'Нет строки выручки кухни':gap.role==='bar'?'Нет строки выручки бара':'Нет общей выручки за день'}`;
    content.append(name,reason);const button=document.createElement('button');button.type='button';button.className='secondaryBtn';button.textContent='Открыть день';
    button.onclick=()=>{
      const details=[...tableBody.querySelectorAll('.detailsRow')].find(el=>el.dataset.restaurant===gap.row.restaurant&&(!el.dataset.date||el.dataset.date===gap.row.dateIso));
      if(!details)return;details.style.display='';details.previousElementSibling.querySelector('.detailBtn').textContent='Скрыть';
      const target=[...details.querySelectorAll('[data-day]')].find(el=>el.dataset.day===gap.row.dateIso)||details;
      target.scrollIntoView({behavior:'smooth',block:'start'});target.classList.add('gapHighlight');setTimeout(()=>target.classList.remove('gapHighlight'),2500);
    };
    item.append(content,button);list.append(item);
  }
  container.append(list);
}
function renderProductivity(rows) {
  const container=document.getElementById('productivity');if(!container)return;
  container.replaceChildren();
  for(const [role,label] of Object.entries(roleLabels)) {
    if(!getCheckedGroups().includes(label))continue;
    const eligible=rows.filter(r=>r.allowedRoles?.includes(role));
    if(!eligible.length) continue;
    const totals={revenue:0,shifts:0,covered:0,missing:0,missingShifts:0};
    eligible.forEach(r=>{const sample=r.coverage?.[role]||productivitySample(r,role);for(const key of Object.keys(totals))totals[key]+=sample[key];});
    const card=document.createElement('article');card.innerHTML=`<span>${label}</span><strong>${totals.shifts>0?formatMoney(totals.revenue/totals.shifts)+" ₽":"—"}</strong><small>${role==='kitchen'?'Выручка кухни':role==='bar'?'Выручка бара':'Общая выручка'} / смены ${label.toLowerCase()}</small><p>В расчёте: ${formatShift(totals.shifts)} смен · ${formatMoney(totals.revenue)} ₽</p>${totals.missing?`<p class="coverageWarning">Не включено: ${totals.missing} предприятие-дней, ${formatShift(totals.missingShifts)} смен без выручки.</p>`:''}`;
    container.append(card);
    if(totals.missing){const button=document.createElement('button');button.className='textBtn';button.type='button';button.textContent='Где пропуски?';button.onclick=()=>{const gaps=document.getElementById('reportGaps');gaps.open=true;gaps.scrollIntoView({behavior:'smooth',block:'start'});};card.append(button);}
  }
  document.getElementById('metricDays').textContent=reportRows?.length?new Set(reportRows.map(r=>r.dateIso)).size:"—";
}
async function showAccess(){
  const status=document.getElementById('accessStatus');
  try{const data=await reportRequest('?resource=access');const list=document.getElementById('accessList');list.replaceChildren();for(const grant of data.grants){const item=document.createElement('div');item.className='grantRow';const text=document.createElement('span');text.textContent=`${grant.email} · ${grant.restaurant} · ${grant.roles.map(r=>roleLabels[r]).join(', ')}`;const remove=document.createElement('button');remove.textContent='Отозвать';remove.className='textBtn';remove.onclick=async()=>{try{await reportRequest('?resource=access',{method:'POST',body:JSON.stringify({...grant,roles:[]})});await showAccess();}catch(e){status.textContent=e.message;}};item.append(text,remove);list.append(item);}}catch(e){status.textContent=e.message;}
}
document.getElementById('accessForm').addEventListener('submit',async(event)=>{event.preventDefault();const status=document.getElementById('accessStatus');try{const roles=[...document.querySelectorAll('[name=accessRole]:checked')].map(i=>i.value);if(!roles.length)throw new Error('Выберите хотя бы один раздел');await reportRequest('?resource=access',{method:'POST',body:JSON.stringify({email:document.getElementById('accessEmail').value,restaurant:document.getElementById('accessRestaurant').value,roles})});status.textContent='Доступ сохранён';await showAccess();}catch(e){status.textContent=e.message;}});
document.getElementById('accessReload').addEventListener('click',showAccess);
document.getElementById('reportView').addEventListener('change',()=>{if(reportRows!==null){lastResultRows=calculateScopedRows();renderTable(lastResultRows);}});
document.querySelectorAll('.groupCheck').forEach(i=>i.addEventListener('change',()=>{if(reportRows!==null){lastResultRows=calculateScopedRows();renderTable(lastResultRows);}}));
let installEvent;
window.addEventListener('beforeinstallprompt',event=>{event.preventDefault();installEvent=event;});
document.getElementById('installApp').onclick=async()=>{if(installEvent){await installEvent.prompt();installEvent=null;}else alert('iPhone: откройте меню «Поделиться» → «На экран Домой». Android: меню браузера → «Установить приложение».');};
if('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js').catch(()=>{});

function reportExportData(rows){
  const roles=['kitchen','hall','bar','delivery'];
  return [["Период / дата","Предприятие","Кухня, смены","Зал, смены","Бар, смены","Доставка, смены","Сохранённая общая выручка","Сохранённая выручка кухни","Сохранённая выручка бара",...roles.flatMap(k=>[`${roleLabels[k]} ₽/смену`,`${roleLabels[k]} смен в расчёте выработки`,`${roleLabels[k]} пропущено предприятие-дней`])],...rows.map(r=>[r.dateIso||`${sabyFromInput.value} — ${sabyToInput.value}`,r.restaurant,...roles.map(k=>r[k]??""),r.revenue??"",r.kitchenRevenue??"",r.barRevenue??"",...roles.flatMap(k=>{const sample=r.coverage?.[k]||productivitySample(r,k);return [sample.shifts>0?sample.revenue/sample.shifts:"",sample.shifts,sample.missing];})])];
}
function exportReportCSV(rows){return reportExportData(rows).map(line=>line.map(value=>`"${String(value).replaceAll('"','""')}"`).join(';')).join('\n');}
function clearScopedSession(){reportRows=[];reportAdmin=false;lastResultRows=[];baseRecords=[];mappedRecords=[];revenueRows=[];staffRestaurantMap.clear();renderTable([]);fillMultiSelect(restaurantSelect,[]);document.getElementById('accessSettings').hidden=true;document.getElementById('forceSyncBtn').hidden=true;document.getElementById('accessList').replaceChildren();}
// Clear financial/personnel data whenever a session changes, including logout.
let scopedUserId=revenueDbSession?.user?.id||null;
if(supabaseClient) supabaseClient.auth.onAuthStateChange((_event,session)=>{if(typeof getSabySession==='function'&&getSabySession())return;if((session?.user?.id||null)!==scopedUserId){clearScopedSession();scopedUserId=session?.user?.id||null;}document.getElementById('sabySettings').hidden=session?.user?.email?.toLowerCase()!=='soltnar@gmail.com';});
