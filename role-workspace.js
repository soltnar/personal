// Presentation only. The existing report endpoint remains the access authority.
let activeReportRole='kitchen';
let compactRows=[];
let loadedReportPeriod=null;
let detailState=null;
const roleDescriptions={operators:'Заказы всей сети с самовывозом / смены операторов',kitchen:'Выручка кухни / смены кухни',hall:'Выручка без доставки / смены зала',bar:'Выручка бара / смены бара',delivery:'Сумма заказов доставки / смены доставки; заказы / смены доставки'};
function refreshIcons(){if(window.feather)window.feather.replace({'stroke-width':1.7,width:18,height:18});}
function shortPeriod(from,to){
  const label=iso=>new Date(`${iso}T12:00:00Z`).toLocaleDateString('ru-RU',{day:'numeric',month:'short',year:'numeric',timeZone:'Europe/Moscow'});
  if(!from||!to)return 'Выбрать период';
  if(from===to)return `${from===yesterdayMoscow()?'Вчера · ':''}${label(from)}`;
  return `${label(from)} — ${label(to)}`;
}
function setLoadedPeriod(from,to){loadedReportPeriod={from,to};syncCompactToolbar();}
function syncCompactToolbar(){
  document.getElementById('periodLabel').textContent=shortPeriod(sabyFromInput.value,sabyToInput.value);
  const filters=document.getElementById('reportFilters');
  document.querySelector('.reportControls')?.classList.toggle('hasReport',!!loadedReportPeriod);
  document.getElementById('toggleFilters').setAttribute('aria-expanded',String(filters.open));
  document.querySelector('.reportControls')?.classList.toggle('periodDirty',!!loadedReportPeriod&&(loadedReportPeriod.from!==sabyFromInput.value||loadedReportPeriod.to!==sabyToInput.value));
  refreshIcons();
}
function reportRoleAllowed(row,role){return row.allowedRoles?.includes(role)&&getCheckedGroups().includes(roleLabels[role]);}
function reportReviewCount(row,role){return Object.keys(roleLabels).filter(key=>(role==='all'||role===key)&&reportRoleAllowed(row,key)).reduce((n,key)=>n+(row.details?.[key]||[]).filter(person=>person.hasConflict).length,0);}
function roleHeadingValues(rows,role){
  const eligible=rows.filter(row=>role==='all'||reportRoleAllowed(row,role));
  const shifts=eligible.reduce((n,row)=>n+(role==='all'?row.total:row[role]||0),0);
  const sample=eligible.reduce((total,row)=>{if(role!=='all'){const value=row.coverage?.[role]||productivitySample(row,role);for(const key of Object.keys(total))total[key]+=value[key];}return total;},{revenue:0,shifts:0,missing:0,missingShifts:0,covered:0});
  return {eligible,shifts,sample};
}
function renderRoleReport(rows){
  compactRows=rows;
  document.getElementById('metricShifts').textContent=rows.length?formatShift(rows.reduce((n,row)=>n+row.total,0)):'—';
  const {known,partial}=summarizeReportRevenue(rows);
  document.getElementById('metricRevenue').textContent=known.length?`${formatMoney(known.reduce((n,row)=>n+row.revenue,0))} ₽${partial?' *':''}`:'—';
  document.getElementById('metricRestaurants').textContent=rows.length?new Set(rows.map(row=>row.restaurant).filter(name=>name!=='Подразделение не определено')).size:'—';
  document.getElementById('metricDays').textContent=new Set(rows.flatMap(row=>row.dailyRows||[row]).map(row=>row.dateIso)).size||'—';
  renderProductivity(rows);renderGaps(rows);refreshWorkspaceReport(rows);
  const available=Object.keys(roleLabels).filter(role=>rows.some(row=>reportRoleAllowed(row,role)));
  if(available.length&&activeReportRole!=='all'&&!available.includes(activeReportRole))activeReportRole=available[0];
  const tabs=document.getElementById('roleTabs');tabs.replaceChildren();
  for(const role of ['all',...available]){
    const button=document.createElement('button');button.type='button';button.textContent=role==='all'?'Все':roleLabels[role];button.setAttribute('aria-pressed',String(role===activeReportRole));button.onclick=()=>{activeReportRole=role;renderRoleReport(compactRows);};tabs.append(button);
  }
  const role=activeReportRole,{eligible,shifts,sample}=roleHeadingValues(rows,role);
  document.getElementById('activeRoleTitle').textContent=role==='all'?'Все разделы':roleLabels[role];
  document.getElementById('roleFormula').textContent=role==='all'?'Смены по выбранным разделам':roleDescriptions[role];
  const revenueTotal=summarizeReportRevenue(eligible).known.reduce((n,row)=>n+row.revenue,0);
  document.getElementById('roleTotals').innerHTML=`<div><strong>${formatShift(shifts)} <small>смен</small></strong><span>${role==='all'?'Во всех выбранных разделах':`Всего: ${roleLabels[role].toLowerCase()}`}</span></div><div><strong>${role==='all'?(eligible.some(row=>Number.isFinite(row.revenue))?`${formatMoney(revenueTotal)} ₽`:'—'):sample.shifts>0?`${formatMoney(sample.revenue/sample.shifts)} <small>₽/смену</small>`:'—'}</strong><span>${role==='all'?'Общая выручка':'Выработка за дни с данными'}</span></div>${['delivery','operators'].includes(role)?deliveryMetrics({dailyRows:eligible.flatMap(row=>row.dailyRows||[row]),[role+'Orders']:eligible.some(row=>Number.isFinite(row[role+'Orders']))?eligible.reduce((n,row)=>n+(row[role+'Orders']||0),0):null},true,role):''}${sample.missing?`<button type="button" id="roleGaps" class="qualityLink">Без выручки: ${sample.missing} предприятие-дней</button>`:''}`;
  document.getElementById('roleGaps')?.addEventListener('click',()=>navigateWorkspace('quality'));
  renderCompactTable();csvBtn.disabled=!rows.length;xlsxBtn.disabled=!rows.length;syncCompactToolbar();
}
function deliveryRowVisible(row,role='delivery'){return row[role]>0||row[role+'Orders']>0||row[role+'Revenue']>0;}
function deliveryOrdersPerShift(row,role='delivery'){
  const days=(row.dailyRows||[row]).filter(day=>Number.isFinite(day[role+'Orders'])&&day[role]>0);
  const shifts=days.reduce((n,day)=>n+day[role],0);
  return shifts?days.reduce((n,day)=>n+day[role+'Orders'],0)/shifts:null;
}
function renderCompactTable(){
  const role=activeReportRole,query=document.getElementById('siteSearch').value.trim().toLocaleLowerCase('ru');
  const rows=compactRows.filter(row=>(role==='all'||reportRoleAllowed(row,role))&&(!['hall','bar'].includes(role)||row[role]>0)&&(!['delivery','operators'].includes(role)||deliveryRowVisible(row,role))&&row.restaurant.toLocaleLowerCase('ru').includes(query));
  const heads=['Предприятие','Смены',role==='kitchen'?'Выручка кухни':role==='bar'?'Выручка бара':role==='delivery'?'Сумма заказов доставки':role==='operators'?'Сумма заказов сети с самовывозом':role==='hall'?'Выручка без доставки':'Общая выручка',role==='all'?'Проверка':'₽/смену',...(['delivery','operators'].includes(role)?['Зак. / смену']:[]),''];
  document.getElementById('resultTable').classList.toggle('deliveryTable',['delivery','operators'].includes(role));
  document.querySelector('#resultTable thead tr').innerHTML=heads.map((label,index)=>`<th scope="col">${label||'<span class="srOnly">Подробнее</span>'}</th>`).join('');
  tableBody.replaceChildren();
  for(const row of rows){
    const index=compactRows.indexOf(row),production=isProductionEnterprise(row),shifts=role==='all'?row.total:row[role],revenue=production?null:row[role==='all'?'revenue':roleRevenue[role]];
    const reviews=reportReviewCount(row,role),sample=role==='all'?null:row.coverage?.[role]||productivitySample(row,role);
    const warning=reviews+(sample?.missing||0),ordersPerShift=['delivery','operators'].includes(role)?deliveryOrdersPerShift(row,role):null;
    const tr=document.createElement('tr');tr.className='compactSiteRow';tr.dataset.siteIndex=String(index);
    const title=escapeHtml(row.restaurant),subtitle=`${row.dateIso?prettyDate(row.dateIso)+' · ':''}${formatShift(shifts)} смен${['delivery','operators'].includes(role)&&Number.isFinite(row[role+'Orders'])?' · '+formatMoney(row[role+'Orders'])+' заказов':''}${production?' · Производство':Number.isFinite(revenue)?' · '+formatMoney(revenue)+' ₽':''}`;
    tr.innerHTML=`<td><button type="button" class="siteOpen" data-site-index="${index}" aria-label="Открыть ${title}"><strong>${title}</strong>${warning?`<span class="qualityDot" title="Отметки требуют проверки: ${reviews}; пропуски выручки: ${sample?.missing||0}" aria-label="Проверить: ${warning}">${warning}</span>`:''}<small>${escapeHtml(subtitle)}</small></button></td><td>${formatShift(shifts)}</td><td>${production?'<span class="productionText">Производство</span>':Number.isFinite(revenue)?formatMoney(revenue)+' ₽':'<span class="missingText">Нет данных</span>'}</td><td class="productivityCell">${role==='all'?(reviews?`<span class="missingText">${reviews}</span>`:'—'):formatProductivity(row,role)}${['delivery','operators'].includes(role)?'<small>₽/смену</small>':''}</td>${['delivery','operators'].includes(role)?`<td class="ordersShiftCell">${Number.isFinite(ordersPerShift)?ordersPerShift.toLocaleString('ru-RU',{maximumFractionDigits:1}):'—'}<small>зак. / смену</small></td>`:''}<td><i data-feather="chevron-right" aria-hidden="true"></i></td>`;
    tableBody.append(tr);
  }
  if(!rows.length)tableBody.innerHTML=`<tr><td colspan="${heads.length}" class="initialEmpty"><strong>${compactRows.length?'Предприятия не найдены':'Откройте отчёт'}</strong><span>${compactRows.length?'Измените поиск или фильтры.':'Выберите период и нажмите «Показать».'}</span></td></tr>`;
  summaryEl.textContent=`${rows.length} ${rows.length===1?'предприятие':'предприятий'}${loadedReportPeriod?' · '+shortPeriod(loadedReportPeriod.from,loadedReportPeriod.to):''}`;
  refreshIcons();
}
function openRoleDetail(row,day=''){
  const allowed=Object.keys(roleLabels).filter(role=>reportRoleAllowed(row,role));
  detailState={row,day,role:allowed.includes(activeReportRole)?activeReportRole:allowed[0]||null};
  renderRoleDetail();const dialog=document.getElementById('siteDialog');if(!dialog.open)dialog.showModal();
}
function deliveryMetrics(row,summary=false,role='delivery'){
  if(summary){
    const daily=row.dailyRows||[row],covered=daily.filter(day=>Number.isFinite(day[role+'Orders'])&&day[role]>0);
    const shifts=covered.reduce((n,day)=>n+day[role],0),orders=covered.reduce((n,day)=>n+day[role+'Orders'],0);
    return `<div class="deliveryMetric"><strong>${Number.isFinite(row[role+'Orders'])?formatMoney(row[role+'Orders']):'—'}</strong><span>${role==='operators'?'Заказы сети с самовывозом':'Заказы на курьеров'}</span></div><div class="deliveryMetric"><strong>${shifts?formatMoney(orders/shifts):'—'}</strong><span>Зак. / смену ${role==='operators'?'операторов':'доставки'}</span></div><p class="deliveryNote">${role==='operators'?'Все заказы сети: доставка и самовывоз.':'Смены доставки включают курьеров, водителей и логистов.'}</p>`;
  }
  if(!Number.isFinite(row[role+'Orders']))return '<p class="hint">Количество курьерских заказов: нет данных.</p>';
  const daily=row.dailyRows||[row],covered=daily.filter(day=>Number.isFinite(day[role+'Orders'])&&day[role]>0);
  const shifts=covered.reduce((n,day)=>n+day[role],0),orders=covered.reduce((n,day)=>n+day[role+'Orders'],0);
  return `<p class="hint">${role==='operators'?'Заказы сети с самовывозом':'Заказы на курьеров'}: <strong>${formatMoney(row[role+'Orders'])}</strong> · Зак. / смену ${role==='operators'?'операторов':'доставки'}: <strong>${shifts?formatMoney(orders/shifts):'—'}</strong>. ${role==='operators'?'Все заказы сети: доставка и самовывоз.':'Смены доставки включают курьеров, водителей и логистов.'}</p>`;
}
function detailMetrics(row,role){
  if(!role)return '';
  if(isProductionEnterprise(row))return `<div class="detailStats"><div><strong>${formatShift(row[role])}</strong><span>Смены</span></div><p class="hint">Производство. Только сотрудники и смены.</p></div>`;
  const field=roleRevenue[role],revenue=row[field],sample=row.coverage?.[role]||productivitySample(row,role);
  return `<div class="detailStats"><div><strong>${formatShift(row[role])}</strong><span>Смены</span></div><div><strong>${Number.isFinite(revenue)?formatMoney(revenue)+' ₽':'—'}</strong><span>${role==='kitchen'?(row.kitchenRevenueSource==='unified'?'Выручка единого склада':'Выручка кухни'):role==='bar'?'Выручка бара':role==='delivery'?'Сумма заказов доставки':role==='operators'?'Сумма заказов сети с самовывозом':role==='hall'?'Выручка без доставки':'Общая выручка'}</span></div><div><strong>${formatProductivity(row,role)}</strong><span>₽/смену</span></div></div>${['delivery','operators'].includes(role)?deliveryMetrics(row,false,role):''}${sample.missing?`<p class="coverageWarning">Выработка не учитывает ${sample.missing} предприятие-дней без выручки (${formatShift(sample.missingShifts)} смен).</p>`:''}`;
}
function renderRoleDetail(){
  if(!detailState)return;
  const {row,role,day}=detailState,days=(row.dailyRows||[row]).slice().sort((a,b)=>a.dateIso.localeCompare(b.dateIso));
  document.getElementById('siteDialogTitle').textContent=row.restaurant;
  document.getElementById('siteDialogPeriod').textContent=loadedReportPeriod?shortPeriod(loadedReportPeriod.from,loadedReportPeriod.to):'';
  const body=document.getElementById('siteDialogBody');body.replaceChildren();
  const tabs=document.createElement('nav');tabs.className='roleTabs';tabs.setAttribute('aria-label','Раздел предприятия');
  for(const key of Object.keys(roleLabels).filter(key=>reportRoleAllowed(row,key))){const button=document.createElement('button');button.type='button';button.textContent=roleLabels[key];button.setAttribute('aria-pressed',String(key===role));button.onclick=()=>{detailState.role=key;renderRoleDetail();};tabs.append(button);}body.append(tabs);
  if(!role){const empty=document.createElement('p');empty.textContent='Выберите доступный раздел в фильтрах.';body.append(empty);return;}
  const control=document.createElement('label');control.className='detailDayControl';control.textContent='День';const select=document.createElement('select');
  if(days.length>1){const option=document.createElement('option');option.value='';option.textContent='Весь период по дням';select.append(option);}
  for(const entry of days){const option=document.createElement('option');option.value=entry.dateIso;option.textContent=prettyDate(entry.dateIso);select.append(option);}select.value=day|| (days.length===1?days[0].dateIso:'');select.onchange=()=>{detailState.day=select.value;renderRoleDetail();};control.append(select);body.append(control);
  const summary=document.createElement('div');summary.innerHTML=detailMetrics(day?days.find(entry=>entry.dateIso===day)||row:row,role);body.append(summary);
  for(const entry of days.filter(entry=>!day||entry.dateIso===day)){
    const section=document.createElement('section');section.className='staffDay';
    section.innerHTML=`<h3>${prettyDate(entry.dateIso)} <small>${formatShift(entry[role])} смен</small></h3>${days.length>1&&!day?detailMetrics(entry,role):''}`;
    const people=entry.details?.[role]||[];
    if(!people.length){const p=document.createElement('p');p.className='hint';p.textContent='Нет сотрудников в выбранном разделе.';section.append(p);}
    else{const list=document.createElement('ul');list.className='staffList';for(const person of people){const item=document.createElement('li');item.classList.toggle('needsReview',!!person.hasConflict);item.innerHTML=`<span><strong>${escapeHtml(person.person)}</strong><small class="staffPosition">${escapeHtml(person.position||'Должность не указана')}</small>${person.hasConflict?`<small>${escapeHtml(person.reviewReason||'Требует проверки')}</small>`:''}</span><span>${formatShift(person.shift)} смен.</span>`;list.append(item);}section.append(list);}
    body.append(section);
  }
  refreshIcons();
}
function openRoleDay(restaurant,date){
  const row=compactRows.find(row=>row.restaurant===restaurant&&(!row.dateIso||row.dateIso===date));
  if(row)openRoleDetail(row,date);
}
function clearRoleWorkspace(){loadedReportPeriod=null;detailState=null;document.getElementById('siteDialog').close();if(typeof resetAccessView==='function')resetAccessView();}
tableBody.addEventListener('click',event=>{const button=event.target.closest('[data-site-index]');if(button&&tableBody.contains(button)){const row=compactRows[Number(button.dataset.siteIndex)];if(row)openRoleDetail(row);}});
document.getElementById('closeSiteDialog').onclick=()=>document.getElementById('siteDialog').close();
document.getElementById('siteSearch').addEventListener('input',renderCompactTable);
document.getElementById('toggleFilters').onclick=()=>{const filters=document.getElementById('reportFilters');filters.open=!filters.open;syncCompactToolbar();};
document.getElementById('reportFilters').addEventListener('toggle',syncCompactToolbar);
document.getElementById('mobileMore').onclick=()=>setWorkspaceMenu(true);
for(const input of [sabyFromInput,sabyToInput])input.addEventListener('change',syncCompactToolbar);
document.querySelectorAll('[data-period]').forEach(button=>button.addEventListener('click',syncCompactToolbar));
loadSabyBtn.addEventListener('click',()=>{document.getElementById('periodPicker').open=false;syncCompactToolbar();});
calcBtn.addEventListener('click',()=>{document.getElementById('reportFilters').open=false;syncCompactToolbar();});
for(const button of [csvBtn,xlsxBtn])button.addEventListener('click',()=>{document.getElementById('reportMore').open=false;});
const moreSummary=document.querySelector('#reportMore>summary');moreSummary.setAttribute('aria-label','Ещё');
const reloadReport=document.createElement('button');reloadReport.type='button';reloadReport.textContent='Обновить отчёт';
document.querySelector('#reportMore .popoverBody').prepend(reloadReport);
reloadReport.onclick=()=>{if(!loadSabyBtn.disabled)loadSabyBtn.click();document.getElementById('reportMore').open=false;};
document.getElementById('reportMore').addEventListener('toggle',()=>{reloadReport.disabled=loadSabyBtn.disabled;});
document.addEventListener('click',event=>{for(const id of ['periodPicker','reportMore']){const details=document.getElementById(id);if(!details.contains(event.target))details.open=false;}});
document.addEventListener('keydown',event=>{if(event.key==='Escape')for(const id of ['periodPicker','reportMore'])document.getElementById(id).open=false;});
// Replace legacy text-symbol icons with the same supplied icon family throughout.
const navIcons={overview:'bar-chart-2',enterprises:'map-pin',quality:'check-square',settings:'settings',users:'users'};
for(const [view,icon] of Object.entries(navIcons)){const span=document.querySelector(`.sidebarNav [data-view="${view}"]>span`);if(span){span.replaceChildren();const i=document.createElement('i');i.dataset.feather=icon;span.append(i);}}
document.querySelector('.brandMark').innerHTML='<img src="personnel-icon-192.png" alt="">';
document.getElementById('menuToggle').innerHTML='<i data-feather="menu" aria-hidden="true"></i>';
syncCompactToolbar();renderRoleReport(lastResultRows||[]);
