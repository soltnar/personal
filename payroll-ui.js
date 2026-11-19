const payrollMonth = document.getElementById('payrollMonth');
const payrollResults = document.getElementById('payrollResults');
const payrollStatus = document.getElementById('payrollStatus');
const payrollImport = document.getElementById('payrollImport');
const payrollNav = document.getElementById('payrollNav');
const payrollMoney = new Intl.NumberFormat('ru-RU', {minimumFractionDigits:2, maximumFractionDigits:2});
let payrollAccess = null;
let payrollGeneration = 0;
let payrollAccessRequest = 0;

function clearPayroll() {
  payrollGeneration++;
  payrollAccessRequest++;
  payrollAccess = null;
  payrollResults.replaceChildren();
  payrollStatus.textContent = '';
  payrollImport.hidden = true;
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
      group.employees.push({id:String(row[id]), name:title, accrual:payrollAmount(row[accrual]), link:String(cell?.l?.Target || row[link] || '')});
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

function renderPayroll(groups) {
  payrollResults.replaceChildren();
  const fmt = (value) => value == null ? '—' : payrollMoney.format(value);
  const total = groups.reduce((sum, group) => sum + group.accrual, 0);
  payrollElement(payrollResults, 'h3', `Начислено: ${fmt(total / 100)} ₽`);
  const summary = payrollTable(payrollResults, ['Подразделение','Начислено, ₽','Смены за месяц','Смены 1–15','Смены с 16','ФОТ / смену, ₽','ФОТ / выручку','Дни с выручкой']);
  for (const group of groups) {
    const row = summary.insertRow();
    [group.name, fmt(group.accrual / 100), group.shifts ?? '—', group.firstHalf ?? '—', group.secondHalf ?? '—',
      fmt(group.shifts ? group.accrual / 100 / group.shifts : null),
      group.revenue > 0 ? fmt(group.accrual / group.revenue) + ' %' : '—', group.revenueDays]
      .forEach(value => payrollElement(row, 'td', value));
  }
  for (const group of groups) {
    const details = payrollElement(payrollResults, 'details', '');
    payrollElement(details, 'summary', `${group.name} · ${group.employees.length} сотрудников`);
    if (group.unallocated) payrollElement(details, 'p', `Разница с итогом подразделения: ${fmt(group.unallocated / 100)} ₽`);
    const employees = payrollTable(details, ['Сотрудник','Начислено, ₽','Смены 1–15','Смены с 16','Начислено / смену, ₽','Карточка']);
    for (const person of group.employees) {
      const row = employees.insertRow();
      const ratio = person.comparable ? person.accrual / 100 / (person.firstHalf + person.secondHalf) : null;
      [person.name, fmt(person.accrual / 100), person.firstHalf ?? '—', person.secondHalf ?? '—', fmt(ratio)]
        .forEach(value => payrollElement(row, 'td', value));
      const cell = payrollElement(row, 'td', '');
      if (person.link) {
        const link = payrollElement(cell, 'a', 'Открыть в СБИС');
        link.href = person.link;
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
      }
    }
  }
}

async function loadPayroll() {
  const generation = ++payrollGeneration;
  payrollResults.replaceChildren();
  if (!payrollAccess) return;
  payrollImport.hidden = !payrollAccess.canImport;
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
const [payrollYear, payrollMonthNumber] = payrollCurrentMonth.split('-').map(Number);
payrollMonth.value = new Date(Date.UTC(payrollYear, payrollMonthNumber - 2, 1)).toISOString().slice(0, 7);
payrollMonth.addEventListener('change', loadPayroll);
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
