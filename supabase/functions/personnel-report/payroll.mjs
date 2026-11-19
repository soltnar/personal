import { canonicalRestaurantName, roleKeys, staffRoleKey, summarizeAttendance } from './engine.mjs';

const nameKey = (value) => String(value ?? '').normalize('NFKC').toLowerCase().replace(/ё/g, 'е').replace(/[-‐‑–—]/g, ' ').trim().replace(/\s+/g, ' ');
const validAmount = (value) => Number.isSafeInteger(value) && Math.abs(value) <= 100_000_000_000;
const validText = (value) => typeof value === 'string' && value.trim().length > 0 && value.length <= 250;

export function validatePayroll(groups) {
  if (!Array.isArray(groups) || !groups.length || groups.length > 200) return null;
  const groupNames = new Set();
  const ids = new Set();
  const result = [];
  for (const group of groups) {
    if (!group || !validText(group.name) || !validAmount(group.accrual) ||
        groupNames.has(group.name) || !Array.isArray(group.employees)) return null;
    groupNames.add(group.name);
    const employees = [];
    for (const entry of group.employees) {
      if (!entry || typeof entry.id !== 'string' || !/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(entry.id) ||
          ids.has(entry.id.toLowerCase()) || !validText(entry.name) || !validAmount(entry.accrual) || ids.size >= 5000) return null;
      ids.add(entry.id.toLowerCase());
      let link = '';
      try {
        const url = new URL(entry.link);
        if (url.protocol === 'https:' && ['online.saby.ru','online.sbis.ru'].includes(url.host) && !url.username && !url.password) link = url.href;
      } catch { /* Some exports omit the link. */ }
      if (entry.position != null && entry.position !== '' && !validText(entry.position)) return null;
      employees.push({id:entry.id.toLowerCase(), name:entry.name, position:entry.position?.trim() || '', accrual:entry.accrual, link});
    }
    result.push({name:group.name, accrual:group.accrual, employees});
  }
  return ids.size ? result : null;
}

function addIdentity(map, key, value) {
  if (!map.has(key)) map.set(key, new Set());
  map.get(key).add(value);
}

function combineBelinskogoPayroll(groups) {
  const point = 'Белинского, 61 Самурай';
  const parts = new Set([point, 'Белинского, 61 доставка Самурай'].map(nameKey));
  const combined = {name:point, accrual:0, employees:[]};
  const result = [];
  for (const group of groups) {
    if (!parts.has(nameKey(group.name))) {
      result.push(group);
      continue;
    }
    if (!result.includes(combined)) result.push(combined);
    combined.accrual += group.accrual;
    combined.employees.push(...group.employees);
  }
  return result;
}

export function buildPayroll(sourceGroups, days, sales, scope, expectedDays, identityDays = []) {
  const groups = combineBelinskogoPayroll(sourceGroups);
  const names = new Map();
  const persons = new Map();
  const payrollNames = new Map();
  const points = new Map();
  for (const group of groups) {
    addIdentity(points, canonicalRestaurantName(group.name), group.name);
    for (const person of group.employees) addIdentity(payrollNames, nameKey(person.name), person.id);
  }
  for (const day of [...days,...identityDays]) for (const event of day.rows || []) {
    const name = nameKey(event.person);
    const id = String(event.personId || '');
    if (!name || !id) continue;
    addIdentity(names, name, id);
    addIdentity(persons, id, name);
  }
  const links = new Map();
  const aliases = new Map();
  for (const [name, ids] of names) {
    const id = ids.values().next().value;
    if (ids.size !== 1 || payrollNames.get(name)?.size !== 1) continue;
    const parts = name.split(' ');
    const compatible = [...persons.get(id)].every(alias => {
      if (names.get(alias).size !== 1) return false;
      if (alias === name) return true;
      if (payrollNames.has(alias)) return false;
      const aliasParts = alias.split(' ');
      return aliasParts.length === 2 && name.startsWith(alias + ' ') ||
        parts.length >= 3 && aliasParts.slice(1).join(' ') === parts.slice(1).join(' ');
    });
    if (!compatible) continue;
    links.set(name,id);
    for (const alias of persons.get(id)) aliases.set(alias,{name,id});
  }
  const normalizedDays = days.map(day => ({
    ...day, rows:(day.rows || []).map(event => {
      const alias=aliases.get(nameKey(event.person));
      if (!alias || event.personId && String(event.personId)!==alias.id) return event;
      return {...event, person:alias.name, personId:alias.id};
    })
  }));
  const pointRows = new Map();
  const employeeRows = new Map();
  for (const row of summarizeAttendance(normalizedDays, sales)) {
    if (!pointRows.has(row.restaurant)) pointRows.set(row.restaurant, []);
    pointRows.get(row.restaurant).push(row);
    for (const [role, details] of Object.entries(row.details)) for (const detail of details) {
      if (!detail.shift) continue;
      const key = nameKey(detail.person);
      if (!employeeRows.has(key)) employeeRows.set(key, []);
      employeeRows.get(key).push({point:row.restaurant, role, day:Number(row.dateIso.slice(-2)), shift:detail.shift});
    }
  }
  const complete = new Set(days.map(day => day.work_date)).size === expectedDays;
  const allowed = new Set(scope.map(canonicalRestaurantName));
  const visible = groups.filter(group => {
    const point = canonicalRestaurantName(group.name);
    return scope.includes('*') || allowed.has(point) && points.get(point).size === 1;
  });
  const payrollGroups = new Map(groups.map(group => [group.name, {...group, employees:[]}]));
  const adjustments = new Map();
  function adjust(point, role, day, amount) {
    if (!adjustments.has(point)) adjustments.set(point,{total:0,firstHalf:0,byRole:{}});
    const value=adjustments.get(point);
    value.total += amount;
    if (day <= 15) value.firstHalf += amount;
    value.byRole[role] = (value.byRole[role] || 0) + amount;
  }
  for (const source of groups) for (const employee of source.employees) {
    const point = canonicalRestaurantName(source.name);
    const technicalStaff = nameKey(source.name) === 'технический персонал';
    const key = nameKey(employee.name);
    const uniquePoint = points.get(point).size === 1;
    const records = complete && links.has(key) && uniquePoint && !technicalStaff ? employeeRows.get(key) || [] : [];
    const {id,name,position='',accrual,link} = employee;
    const category = technicalStaff || /отзыв/i.test(position) ? 'other' : staffRoleKey(position,point) || 'other';
    let firstHalf = records.length ? 0 : null;
    let secondHalf = firstHalf;
    for (const item of records) {
      if (item.day <= 15) firstHalf += item.shift;
      else secondHalf += item.shift;
      if (item.point === point) continue;
      adjust(item.point,item.role,item.day,-item.shift);
      adjust(point,category === 'other' ? item.role : category,item.day,item.shift);
    }
    payrollGroups.get(source.name).employees.push({id,name,position,accrual,link,category,firstHalf,secondHalf,comparable:records.length > 0});
  }
  return visible.map(source => {
    const group = payrollGroups.get(source.name);
    const point = canonicalRestaurantName(group.name);
    const rows = points.get(point).size === 1 ? pointRows.get(point) || [] : [];
    const change = points.get(point).size === 1 ? adjustments.get(point) : null;
    const canCount = complete && (change || rows.some(row => row.total > 0 || Object.values(row.details).some(items => items.length)));
    const shifts = canCount ? rows.reduce((sum, row) => sum + row.total, change?.total || 0) : null;
    const firstHalf = canCount ? rows.filter(row => Number(row.dateIso.slice(-2)) <= 15)
      .reduce((sum, row) => sum + row.total, change?.firstHalf || 0) : null;
    const salesRows = rows.filter(row => Number.isFinite(row.revenue));
    const observedRevenue = salesRows.length ? salesRows.reduce((sum, row) => sum + row.revenue, 0) : null;
    const revenue = salesRows.length === expectedDays ? observedRevenue : null;
    const knownRevenueDays = new Set(salesRows.map(row => Number(row.dateIso.slice(-2))));
    const missingRevenueDays = Array.from({length:expectedDays}, (_, index) => index + 1)
      .filter(day => !knownRevenueDays.has(day));
    const categoryShifts = Object.fromEntries(roleKeys.map(role => [
      role, canCount ? rows.reduce((sum, row) => sum + row[role], change?.byRole[role] || 0) : null,
    ]));
    const categoryTotals = {kitchen:0,hall:0,bar:0,delivery:0,operators:0,other:0};
    for (const employee of group.employees) categoryTotals[employee.category] += employee.accrual;
    return {
      name:group.name, accrual:group.accrual, shifts, firstHalf, secondHalf:canCount ? shifts - firstHalf : null, revenue, observedRevenue, revenueDays:salesRows.length,
      unallocated:group.accrual - Object.values(categoryTotals).reduce((sum, value) => sum + value, 0),
      categoryTotals, categoryShifts, missingRevenueDays, expectedDays, employees:group.employees,
    };
  });
}
