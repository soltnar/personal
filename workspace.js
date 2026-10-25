// Navigation is presentation only. Report APIs remain the authority for access.
const workspaceViews = {
  overview: ['Обзор', 'ПОКАЗАТЕЛИ СЕТИ', 'Смены, выручка и выработка — всё главное за выбранный период.'],
  enterprises: ['По предприятиям', 'ДЕТАЛЬНЫЙ ОТЧЁТ', 'Выручка, смены и выработка по предприятиям.'],
  quality: ['Проверка данных', 'ПОЛНОТА И КАЧЕСТВО', 'Найдите пропуски и неоднозначные отметки, прежде чем делать выводы.'],
  settings: ['Настройки', 'ДАННЫЕ И ПРИЛОЖЕНИЕ', 'Синхронизация проходной, правила расчёта и установка приложения.'],
  users: ['Пользователи и доступ', 'УПРАВЛЕНИЕ', 'Подключение Saby и права сотрудников по предприятиям и разделам.'],
  account: ['Мой профиль', 'АККАУНТ', 'Кто вошёл в приложение и какие данные ему доступны.']
};
let workspaceIdentity = null;
function workspaceIsOwner() {
  return !getSabySession() && revenueDbSession?.user?.email?.toLowerCase() === 'soltnar@gmail.com';
}
function setWorkspaceMenu(open) {
  document.body.classList.toggle('menuOpen', open);
  document.getElementById('menuToggle').setAttribute('aria-expanded', String(open));
  document.getElementById('navBackdrop').hidden = !open;
  document.querySelector('.app').inert = open && window.innerWidth<=900;
  if(open) document.querySelector('.sidebarNav a[aria-current="page"]')?.focus();
}
function showWorkspace(view, focus = false) {
  if(!Object.hasOwn(workspaceViews, view)) view = 'overview';
  if(view === 'users' && !workspaceIsOwner()) view = 'account';
  const [title, eyebrow, description] = workspaceViews[view];
  document.querySelectorAll('[data-pages]').forEach(el => {
    const permitted = (!el.hasAttribute('data-admin-only') || workspaceIsOwner()) &&
      (!el.hasAttribute('data-report-admin') || (workspaceIsOwner() && reportAdmin));
    el.hidden = !el.dataset.pages.split(' ').includes(view) || !permitted;
  });
  document.getElementById('usersEmpty').hidden = view !== 'users' || workspaceIsOwner();
  document.querySelectorAll('.sidebarNav [data-view]').forEach(link => {
    if(link.dataset.view === view) link.setAttribute('aria-current', 'page');
    else link.removeAttribute('aria-current');
  });
  document.getElementById('workspaceTitle').textContent = title;
  document.getElementById('breadcrumbCurrent').textContent = title;
  document.getElementById('workspaceEyebrow').textContent = eyebrow;
  document.getElementById('workspaceDescription').textContent = description;
  document.title = `${title} · Персонал и выручка`;
  document.getElementById('usersNav').hidden = !workspaceIsOwner();
  document.getElementById('syncFrom').value=sabyFromInput.value;
  document.getElementById('syncTo').value=sabyToInput.value;
  document.querySelectorAll('[data-mobile-view]').forEach(link=>{if(link.dataset.mobileView===view)link.setAttribute('aria-current','page');else link.removeAttribute('aria-current');});
  if(view==='users'&&workspaceIsOwner()&&typeof ensureAccessLoaded==='function')ensureAccessLoaded();
  if(typeof syncCompactToolbar==='function')syncCompactToolbar();
  setWorkspaceMenu(false);
  if(focus) { document.getElementById('workspaceTitle').focus({preventScroll:true}); window.scrollTo({top:0,behavior:'instant'}); }
}
function navigateWorkspace(view) {
  if(location.hash === `#${view}`) showWorkspace(view, true);
  else location.hash = view;
}
function syncWorkspaceProfile() {
  const saby = getSabySession(), user = revenueDbSession?.user;
  const identity = saby ? `saby:${saby.subject}` : user?.id || null;
  const changed = workspaceIdentity !== null && workspaceIdentity !== identity;
  workspaceIdentity = identity;
  if(changed) clearScopedSession();
  const signedIn = !!identity, owner = workspaceIsOwner();
  const name = saby?.name || user?.user_metadata?.full_name || user?.user_metadata?.name || user?.email || 'Гость';
  const initials = signedIn ? name.split(/\s+/).filter(Boolean).slice(0,2).map(part => part[0]).join('').toUpperCase() : '?';
  const provider = saby ? 'Saby' : signedIn ? 'Google' : '—';
  document.querySelectorAll('[data-profile-name]').forEach(el => {el.textContent = name; el.title = name;});
  document.querySelectorAll('[data-profile-avatar]').forEach(el => {el.textContent = initials;});
  document.querySelectorAll('[data-profile-caption]').forEach(el => {el.textContent = signedIn ? `${provider} · ${owner?'Владелец':'Сотрудник'}` : 'Войдите в аккаунт';});
  document.getElementById('accountName').textContent = signedIn ? name : 'Вы не вошли';
  document.getElementById('accountIdentifier').textContent = saby ? `Saby ID: ${saby.subject}` : user?.email || 'Войдите, чтобы открыть доступные вам предприятия.';
  document.getElementById('accountProvider').textContent = provider;
  document.getElementById('accountRole').textContent = owner ? 'Владелец · управление доступом' : signedIn ? 'Сотрудник · назначенный доступ' : 'Гость';
  let scope = 'Данные закрыты до входа';
  if(owner) scope = 'Все предприятия и разделы. Управление пользователями и синхронизацией.';
  else if(signedIn && reportRows?.length) {
    const restaurants = [...new Set(reportRows.map(row=>row.restaurant))];
    const roles = [...new Set(reportRows.flatMap(row=>row.allowedRoles||[]))].map(role=>roleLabels[role]);
    scope = `В текущем отчёте: ${restaurants.join('; ')}. Разделы: ${roles.join(', ')}.`;
  } else if(signedIn) scope = 'Назначения проверяются сервером при открытии отчёта. Если доступ ещё не выдан, обратитесь к администратору.';
  document.getElementById('accountScope').textContent = scope;
  document.getElementById('reportSignInLink').hidden = signedIn;
  document.getElementById('forceSyncBtn').hidden = !owner;
  document.getElementById('syncAccessHint').hidden = owner;
  document.getElementById('sabySetupHint').hidden = signedIn || !document.getElementById('sabyLoginBtn').disabled;
  showWorkspace(location.hash.slice(1) || 'overview');
}
function openWorkspaceDay(restaurant, date) {
  if(typeof openRoleDay==='function'&&reportRows!==null){openRoleDay(restaurant,date);return;}
  navigateWorkspace('enterprises');
  // Reveal now as hashchange is asynchronous; the row must be visible before scrolling.
  showWorkspace('enterprises');
  const detail = [...tableBody.querySelectorAll('.detailsRow')].find(el=>el.dataset.restaurant===restaurant && (!el.dataset.date || el.dataset.date===date));
  if(!detail) return;
  detail.style.display = '';
  detail.previousElementSibling.querySelector('.detailBtn').textContent = 'Скрыть';
  const target = [...detail.querySelectorAll('[data-day]')].find(el=>el.dataset.day===date) || detail;
  requestAnimationFrame(()=>{target.scrollIntoView({behavior:'smooth',block:'start'});target.classList.add('gapHighlight');setTimeout(()=>target.classList.remove('gapHighlight'),2500);});
}
function refreshWorkspaceReport(rows) {
  const daily = rows.flatMap(row=>row.dailyRows||[row]);
  const checks = [],informational = [];
  let gaps = 0;
  for(const row of daily) for(const [role,label] of Object.entries(roleLabels)) {
    if(!row.allowedRoles?.includes(role) || !getCheckedGroups().includes(label)) continue;
    if((row.coverage?.[role] || productivitySample(row,role)).missing) gaps++;
    for(const person of row.details?.[role] || []) if(person.hasConflict) checks.push({row, person, label});
  }
  for(const row of daily)for(const person of row.unclassified||[])(person.excludedFromDelivery?informational:checks).push({row,person,label:person.excludedFromDelivery?'Вне расчётов доставки':'Без категории'});
  const target = document.getElementById('attendanceReviews');
  target.replaceChildren();
  const heading = document.createElement('h3');heading.className = 'panelTitle';heading.textContent = `Отметки сотрудников · ${checks.length}`;target.append(heading);
  const explanation = document.createElement('p');explanation.className = 'hint';explanation.textContent = rows.length ? (checks.length ? 'Причины проверки отмечены рядом с сотрудниками в раскрытии соответствующего дня.' : 'Неоднозначных отметок в текущей выборке нет.') : 'Сначала откройте отчёт за нужный период.';target.append(explanation);
  const list = document.createElement('div');list.className = 'gapList reviewList';
  const info=document.createElement('details');info.className='excludedAttendance';
  const infoTitle=document.createElement('summary');infoTitle.textContent='Водитель газели — отметки вне расчётов';
  const infoList=document.createElement('div');infoList.className='gapList';info.append(infoTitle,infoList);
  for(const {row,person,label} of [...checks,...informational]) {
    const item = document.createElement('article'), content = document.createElement('div'), title = document.createElement('strong'), description = document.createElement('p'), button = document.createElement('button');
    title.textContent = `${row.restaurant} · ${prettyDate(row.dateIso)} · ${label}`;
    description.textContent = `${person.person || 'Сотрудник'}${person.position ? ' · '+person.position : ''}${person.marks?.length ? ' · Отметки: '+person.marks.map(mark=>mark.dateTime+' '+(Number(mark.actionType)===1?'вход':'выход')+(mark.location?' ('+mark.location+')':'')).join(', ') : ''}${person.reviewReason ? ' · '+person.reviewReason : ''}`;
    button.type='button';button.className='secondaryBtn';button.textContent='Открыть день';button.onclick=()=>openWorkspaceDay(row.restaurant,row.dateIso);
    content.append(title,description);item.append(content,button);(person.excludedFromDelivery?infoList:list).append(item);
  }
  target.append(list);
  if(informational.length)target.append(info);
  const badge=document.getElementById('qualityBadge');badge.textContent=String(gaps+checks.length);badge.hidden=!(gaps+checks.length);
  document.getElementById('overviewEmpty').hidden=rows.length>0;
  document.getElementById('filterSelectionLabel').textContent=rows.length ? `${new Set(daily.map(r=>r.restaurant)).size} предприятий · ${getCheckedGroups().join(', ')}` : 'Настроить выборку';
  if(rows.length) document.getElementById('metricDays').textContent=new Set(daily.map(r=>r.dateIso)).size;
  syncWorkspaceProfile();
}
document.getElementById('menuToggle').onclick=()=>setWorkspaceMenu(!document.body.classList.contains('menuOpen'));
document.getElementById('navBackdrop').onclick=()=>setWorkspaceMenu(false);
document.querySelectorAll('#appSidebar a').forEach(link=>link.addEventListener('click',event=>{event.preventDefault();navigateWorkspace(link.hash.slice(1));}));
document.addEventListener('keydown',event=>{
  if(!document.body.classList.contains('menuOpen'))return;
  if(event.key==='Escape'){setWorkspaceMenu(false);document.getElementById('menuToggle').focus();}
  if(event.key==='Tab'){
    const links=[...document.querySelectorAll('#appSidebar a')].filter(el=>!el.hidden);
    const first=links[0],last=links[links.length-1];
    if(event.shiftKey&&document.activeElement===first){event.preventDefault();last.focus();}
    else if(!event.shiftKey&&document.activeElement===last){event.preventDefault();first.focus();}
  }
});
window.addEventListener('hashchange',()=>showWorkspace(location.hash.slice(1),true));
window.addEventListener('resize',()=>{if(window.innerWidth>900)setWorkspaceMenu(false);});
for(const [syncId,input] of [['syncFrom',sabyFromInput],['syncTo',sabyToInput]])document.getElementById(syncId).addEventListener('change',event=>{input.value=event.target.value;updateRevenueDbButtons();});
syncWorkspaceProfile();
refreshWorkspaceReport(lastResultRows || []);
