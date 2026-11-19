const payrollMonth = document.getElementById('payrollMonth');
const payrollYear = document.getElementById('payrollYear');
const payrollResults = document.getElementById('payrollResults');
const payrollStatus = document.getElementById('payrollStatus');
const payrollImport = document.getElementById('payrollImport');
const payrollUpload = document.getElementById('payrollUpload');
const payrollTabs = document.getElementById('payrollTabs');
const payrollHeading = document.getElementById('payrollHeading');
const payrollSearch = document.getElementById('payrollSearch');
const payrollNav = document.getElementById('payrollNav');
const payrollCategories = {kitchen:'Кухня',hall:'Зал',bar:'Бар',delivery:'Доставка',operators:'Операторы',other:'Прочие должности'};
const payrollMoney = new Intl.NumberFormat('ru-RU', {minimumFractionDigits:2, maximumFractionDigits:2});
let payrollAccess = null;
let payrollGeneration = 0;
let payrollAccessRequest = 0;
let payrollRows = [];
let payrollCategory = 'all';

function clearPayroll() {
  payrollGeneration++;
  payrollAccessRequest++;
  payrollAccess = null;
  clearPayrollResults();
  payrollStatus.textContent = '';
  payrollUpload.hidden = true;
  payrollUpload.open = false;
  payrollSearch.value = '';
  payrollCategory = 'all';
  payrollNav.hidden = true;
  document.getElementById('payrollFile').value = '';
}

async function syncPayrollAccess(identity) {
  const request = ++payrollAccessRequest;
  payrollAccess = null;
  payrollNav.hidden = true;
  if (!identity) {
    payrollStatus.textContent = 'Войдите для просмотра ФОТ.';
    return;
  }
  try {
    const access = await reportRequest('?resource=payroll&action=access');
    if (request !== payrollAccessRequest) return;
    payrollAccess = access;
    payrollNav.hidden = false;
    if (location.hash === '#payroll') await loadPayroll();
  } catch {
    if (request === payrollAccessRequest) payrollStatus.textContent = 'Доступ к ФОТ не назначен или сервис недоступен.';
  }
}

function payrollAmount(value) {
  if (value == null || value === '') return 0;
  if (typeof value !== 'number' || !Number.isFinite(value)) throw new Error('В ведомости есть некорректная сумма.');
  return Math.round(value * 100);
}

function parsePayrollWorkbook(book) {
  const sheet = book.Sheets[book.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, {header:1, raw:true, blankrows:true, range:0});
  const header = rows[0] || [];
  const [name, accrual, id, link] = ['Подразделение/ФИО','Начисления','Идентификатор сотрудника','Ссылка на сотрудника'].map(title => header.indexOf(title));
  if ([name, accrual, id].some(column => column < 0) || !sheet['!rows']) {
    throw new Error('Загрузите Excel из отчёта СБИС «Расчётная ведомость» по сотрудникам.');
  }
  const [position, organization] = ['Должность','Организация'].map(title => header.indexOf(title));
  const groups = [];
  let group;
  for (const [index, row] of rows.entries()) {
    if (!index || !row.some(value => value != null && value !== '')) continue;
    const level = sheet['!rows'][index]?.level || 0;
    const title = String(row[name] ?? '').trim();
    if (!title) throw new Error(`Не указано подразделение или ФИО в строке ${index + 1}.`);
    if (!level && !row[id]) {
      group = {name:title, accrual:payrollAmount(row[accrual]), employees:[]};
      groups.push(group);
    } else if (row[id] && level && group) {
      const cell = link >= 0 && sheet[XLSX.utils.encode_cell({r:index,c:link})];
      group.employees.push({id:String(row[id]), name:title, position:String(row[position] ?? '').trim(), accrual:payrollAmount(row[accrual]), link:String(cell?.l?.Target || row[link] || '')});
    } else if (!level || !group || row[position] || row[organization]) {
      throw new Error(`Не удалось определить подразделение в строке ${index + 1}.`);
    }
  }
  if (!groups.some(item => item.employees.length)) throw new Error('В ведомости нет строк сотрудников.');
  return groups;
}

function payrollElement(parent, tag, text) {
  const element = document.createElement(tag);
  element.textContent = text;
  parent.append(element);
  return element;
}

function payrollTable(parent, titles) {
  const wrap = payrollElement(parent, 'div', '');
  wrap.className = 'tableWrap';
  const table = payrollElement(wrap, 'table', '');
  const head = table.createTHead().insertRow();
  titles.forEach(title => payrollElement(head, 'th', title).scope = 'col');
  return table.createTBody();
}

const payrollNumber = new Intl.NumberFormat('ru-RU', {maximumFractionDigits:1});
const payrollRubles = (value) => value == null ? '—' : payrollMoney.format(value) + ' ₽';
const payrollShifts = (value) => value == null ? '—' : payrollNumber.format(value);

function clearPayrollResults() {
  payrollRows = [];
  payrollResults.replaceChildren();
  payrollTabs.hidden = true;
  payrollHeading.hidden = true;
}

function payrollCoverage(group) {
  if (group.revenueDays === 0) return 'Выручка не найдена';
  if (group.expectedDays && group.revenueDays < group.expectedDays) {
    const month = Number(payrollMonth.value.slice(5));
    const dates = group.missingRevenueDays.map(day => `${String(day).padStart(2,'0')}.${String(month).padStart(2,'0')}`);
    return `Выручка за ${group.revenueDays} из ${group.expectedDays} дней. Нет данных: ${dates.join(', ')}.`;
  }
  if (!((group.observedRevenue ?? group.revenue) > 0)) return 'Выручка равна нулю';
  return '';
}

function fillPayrollDetails(container, group, people, category) {
  const notes = [];
  if (category === 'all') {
    notes.push(`Смены за 1–15: ${payrollShifts(group.firstHalf)} · с 16 числа: ${payrollShifts(group.secondHalf)}`);
    if (group.unallocated) notes.push(`Не распределено по сотрудникам: ${payrollRubles(group.unallocated / 100)}`);
  }
  const coverage = payrollCoverage(group);
  if (coverage) notes.push(coverage);
  if (group.observedRevenue != null) {
    notes.push(`${coverage ? 'Выручка за дни с данными' : 'Выручка за месяц'}: ${payrollRubles(group.observedRevenue)}`);
  }
  if (category === 'all' && group.categoryTotals) {
    const roles = Object.entries(payrollCategories).filter(([key]) => group.categoryTotals[key]);
    if (roles.length) {
      const breakdown = payrollElement(container,'div','');
      breakdown.className = 'payrollBreakdown';
      for (const [key,label] of roles) {
        const item = payrollElement(breakdown,'span','');
        payrollElement(item,'small',label);
        payrollElement(item,'strong',payrollRubles(group.categoryTotals[key] / 100));
      }
    }
  }
  for (const note of notes) payrollElement(container,'p',note).className = 'hint';
  if (!people.length) {
    payrollElement(container,'p','В ведомости нет сотрудников этой категории.').className = 'hint';
    return;
  }
  const employees = payrollTable(container, ['Сотрудник','Должность','Категория','Начислено','Смены 1–15','Смены с 16','Начислено / смену','']);
  const categoryOrder = Object.keys(payrollCategories);
  for (const person of [...people].sort((a,b) => categoryOrder.indexOf(a.category)-categoryOrder.indexOf(b.category) || a.name.localeCompare(b.name,'ru'))) {
    const row = employees.insertRow();
    const ratio = person.comparable ? person.accrual / 100 / (person.firstHalf + person.secondHalf) : null;
    [person.name,person.position || '—',payrollCategories[person.category] || payrollCategories.other,payrollRubles(person.accrual / 100),
      payrollShifts(person.firstHalf),payrollShifts(person.secondHalf),payrollRubles(ratio)]
      .forEach(value => payrollElement(row,'td',value));
    const cell = payrollElement(row,'td','');
    if (person.link) {
      const link = payrollElement(cell,'a','В СБИС');
      link.href = person.link;
      link.target = '_blank';
      link.rel = 'noopener noreferrer';
    }
  }
}

function renderPayrollRows() {
  payrollResults.replaceChildren();
  const category = payrollCategory;
  const all = category === 'all';
  document.getElementById('payrollCategoryTitle').textContent = all ? 'Все категории' : payrollCategories[category];
  const search = payrollSearch.value.trim().toLocaleLowerCase('ru').replace(/ё/g,'е');
  const rows = payrollRows.filter(group => group.name.toLocaleLowerCase('ru').replace(/ё/g,'е').includes(search))
    .map(group => ({
      group,
      accrual:all ? group.accrual : group.categoryTotals?.[category] ?? 0,
      shifts:all ? group.shifts : group.categoryShifts?.[category] ?? null,
      people:all ? group.employees : group.employees.filter(person => person.category === category),
    })).filter(row => all || row.people.length || row.shifts > 0);
  if (!rows.length) {
    payrollElement(payrollResults,'p',search ? 'Предприятия не найдены.' : 'За этот месяц нет данных по категории.').className = 'hint';
    return;
  }

  const totals = payrollElement(payrollResults,'div','');
  totals.className = 'roleTotals payrollTotals';
  const shiftRows = rows.filter(row => row.shifts != null);
  const shiftTotal = shiftRows.reduce((sum,row) => sum + row.shifts,0);
  const revenueFor = (group) => group.observedRevenue ?? group.revenue;
  const comparable = rows.filter(row => revenueFor(row.group) > 0);
  const revenueTotal = comparable.reduce((sum,row) => sum + revenueFor(row.group),0);
  const accrualTotal = rows.reduce((sum,row) => sum + row.accrual,0);
  const percent = (amount,revenue) => revenue > 0 ? payrollMoney.format(amount / revenue) + ' %' : '—';
  for (const [label,value] of [
    ['Начислено',payrollRubles(accrualTotal / 100)],
    [shiftRows.length < rows.length ? 'Смены по точкам с данными' : 'Учтённые смены',payrollShifts(shiftRows.length ? shiftTotal : null)],
    [all ? 'ФОТ / выручка · по точкам с выручкой' : 'ФОТ категории / выручка её точек',percent(comparable.reduce((sum,row) => sum + row.accrual,0),revenueTotal)]
  ]) {
    const item = payrollElement(totals,'div','');
    payrollElement(item,'strong',value);
    payrollElement(item,'span',label);
  }
  const body = payrollTable(payrollResults,['Предприятие','Начислено','Смены','ФОТ / смену','Выручка точки','ФОТ / выручка']);
  body.closest('table').className = 'payrollSites';
  for (const {group,accrual,shifts,people} of rows) {
    const row = body.insertRow();
    const name = payrollElement(row,'td','');
    const button = payrollElement(name,'button','');
    button.className = 'siteOpen';
    button.type = 'button';
    button.setAttribute('aria-expanded','false');
    payrollElement(button,'strong',group.name);
    const coverage = payrollCoverage(group);
    if (coverage && group.revenueDays > 0) {
      const warning = payrollElement(button,'span','!');
      warning.className = 'qualityDot';
      warning.title = coverage;
      warning.setAttribute('aria-label',coverage);
    }
    payrollElement(button,'small',`${payrollRubles(accrual / 100)} · ${payrollShifts(shifts)} смен`);
    const ratio = shifts > 0 ? accrual / 100 / shifts : null;
    [payrollRubles(accrual / 100),payrollShifts(shifts),payrollRubles(ratio)]
      .forEach(value => payrollElement(row,'td',value));
    const revenueCell = payrollElement(row,'td',payrollRubles(revenueFor(group)));
    if (group.revenueDays > 0 && group.expectedDays && group.revenueDays < group.expectedDays) {
      payrollElement(revenueCell,'small',`${group.revenueDays} из ${group.expectedDays} дней`).className = 'payrollCoverage';
    }
    const percentCell = payrollElement(row,'td',percent(accrual,revenueFor(group)));
    if (group.revenue == null && group.revenueDays > 0 && group.expectedDays) {
      payrollElement(percentCell,'small',`Выручка за ${group.revenueDays} из ${group.expectedDays} дней`).className = 'payrollCoverage';
    }
    const disclosure = body.insertRow();
    disclosure.hidden = true;
    disclosure.className = 'payrollDisclosure';
    const cell = disclosure.insertCell();
    cell.colSpan = 6;
    button.onclick = () => {
      disclosure.hidden = !disclosure.hidden;
      button.setAttribute('aria-expanded',String(!disclosure.hidden));
      if (!disclosure.hidden && !cell.hasChildNodes()) fillPayrollDetails(cell,group,people,category);
    };
  }
}

function renderPayroll(groups) {
  payrollRows = groups;
  payrollTabs.replaceChildren();
  payrollTabs.hidden = false;
  payrollHeading.hidden = false;
  for (const [key,label] of [['all','Все'],...Object.entries(payrollCategories)]) {
    const tab = payrollElement(payrollTabs,'button',label);
    tab.type = 'button';
    tab.setAttribute('aria-pressed',String(key===payrollCategory));
    tab.onclick = () => {
      payrollCategory = key;
      for (const button of payrollTabs.children) button.setAttribute('aria-pressed',String(button===tab));
      renderPayrollRows();
    };
  }
  renderPayrollRows();
}
async function loadPayroll() {
  const generation = ++payrollGeneration;
  clearPayrollResults();
  if (!payrollAccess) return;
  payrollUpload.hidden = !payrollAccess.canImport;
  payrollStatus.textContent = 'Загружаем ФОТ…';
  try {
    const result = await reportRequest(`?resource=payroll&month=${encodeURIComponent(payrollMonth.value)}`);
    if (generation !== payrollGeneration) return;
    if (!result.groups.length) {
      payrollStatus.textContent = 'За этот месяц нет доступной ведомости.';
      return;
    }
    renderPayroll(result.groups);
    payrollStatus.textContent = payrollMonth.value >= payrollCurrentMonth
      ? 'Текущий месяц: предварительные начисления по последней выгрузке СБИС.'
      : '';
  } catch (error) {
    if (generation === payrollGeneration) payrollStatus.textContent = error.message;
  }
}

const payrollCurrentMonth = new Date().toLocaleDateString('sv-SE', {timeZone:'Europe/Moscow'}).slice(0, 7);
const [currentYear, currentMonth] = payrollCurrentMonth.split('-').map(Number);
const previousMonth = new Date(Date.UTC(currentYear, currentMonth - 2, 1));
const payrollMonthNames = new Intl.DateTimeFormat('ru-RU', {month:'long',timeZone:'UTC'});

function fillPayrollMonths(monthNumber) {
  payrollMonth.replaceChildren();
  for (let month = 1; month <= 12; month++) {
    const label = payrollMonthNames.format(new Date(Date.UTC(Number(payrollYear.value),month - 1,1)));
    const option = payrollElement(payrollMonth,'option',label[0].toLocaleUpperCase('ru') + label.slice(1));
    option.value = `${payrollYear.value}-${String(month).padStart(2,'0')}`;
  }
  payrollMonth.selectedIndex = monthNumber - 1;
}

for (let year = currentYear; year >= 2000; year--) {
  payrollElement(payrollYear,'option',String(year)).value = String(year);
}
payrollYear.value = String(previousMonth.getUTCFullYear());
fillPayrollMonths(previousMonth.getUTCMonth() + 1);
payrollYear.addEventListener('change', () => {
  fillPayrollMonths(payrollMonth.selectedIndex + 1);
  loadPayroll();
});
payrollMonth.addEventListener('change', loadPayroll);
payrollSearch.addEventListener('input', renderPayrollRows);
payrollImport.addEventListener('submit', async (event) => {
  event.preventDefault();
  if (!payrollAccess?.canImport) return;
  const generation = ++payrollGeneration;
  const file = document.getElementById('payrollFile');
  const submit = payrollImport.querySelector('button');
  const month = payrollMonth.value;
  submit.disabled = true;
  try {
    const session = await getReportSession();
    const source = file.files[0];
    if (!source || source.size > 5_000_000) throw new Error('Выберите файл Excel размером до 5 МБ.');
    const knownMonth = /за\s+(январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь|декабрь)['’](\d{2})/i.exec(source.name);
    if (knownMonth) {
      const months = 'январь февраль март апрель май июнь июль август сентябрь октябрь ноябрь декабрь'.split(' ');
      const fileMonth = '20' + knownMonth[2] + '-' + String(months.indexOf(knownMonth[1].toLowerCase()) + 1).padStart(2, '0');
      if (fileMonth !== month) throw new Error('Месяц в имени ведомости не совпадает с выбранным.');
    }
    const groups = parsePayrollWorkbook(XLSX.read(await source.arrayBuffer(), {type:'array',cellStyles:true}));
    if (generation !== payrollGeneration || !session || session.user.id !== (await getReportSession())?.user.id) return;
    await reportRequest(`?resource=payroll&month=${encodeURIComponent(month)}`, {method:'POST', body:JSON.stringify({groups})});
    if (generation === payrollGeneration) {
      file.value = '';
      await loadPayroll();
    }
  } catch (error) {
    if (generation === payrollGeneration) payrollStatus.textContent = error.message;
  } finally {
    submit.disabled = false;
  }
});
