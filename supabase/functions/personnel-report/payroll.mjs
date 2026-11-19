import { canonicalRestaurantName, summarizeAttendance } from './engine.mjs';

const nameKey = (value) => String(value ?? '').normalize('NFKC').toLowerCase().replace(/ё/g, 'е').trim().replace(/\s+/g, ' ');
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
      employees.push({id:entry.id.toLowerCase(), name:entry.name, accrual:entry.accrual, link});
    }
    result.push({name:group.name, accrual:group.accrual, employees});
  }
  return ids.size ? result : null;
}

function addIdentity(map, key, value) {
  if (!map.has(key)) map.set(key, new Set());
  map.get(key).add(value);
}

export function buildPayroll(groups, days, sales, scope, expectedDays) {
  const names = new Map();
  const persons = new Map();
  const payrollNames = new Map();
  const points = new Map();
  for (const group of groups) {
    addIdentity(points, canonicalRestaurantName(group.name), group.name);
    for (const person of group.employees) addIdentity(payrollNames, nameKey(person.name), person.id);
  }
  for (const day of days) for (const event of day.rows || []) {
    const name = nameKey(event.person);
    const id = String(event.personId || '');
    if (!name || !id) continue;
    addIdentity(names, name, id);
    addIdentity(persons, id, name);
  }
  const links = new Map();
  for (const [name, ids] of names) {
    const id = ids.values().next().value;
    if (ids.size === 1 && persons.get(id).size === 1 && payrollNames.get(name)?.size === 1) links.set(name, id);
  }
  const normalizedDays = days.map(day => ({
    ...day, rows:(day.rows || []).map(event => event.personId ? event : {...event, personId:links.get(nameKey(event.person)) || ''})
  }));
  const pointRows = new Map();
  const employeeRows = new Map();
  for (const row of summarizeAttendance(normalizedDays, sales)) {
    if (!pointRows.has(row.restaurant)) pointRows.set(row.restaurant, []);
    pointRows.get(row.restaurant).push(row);
    for (const detail of Object.values(row.details).flat()) {
      if (!detail.shift) continue;
      const key = nameKey(detail.person);
      if (!employeeRows.has(key)) employeeRows.set(key, []);
      employeeRows.get(key).push({point:row.restaurant, day:Number(row.dateIso.slice(-2)), shift:detail.shift});
    }
  }
  const complete = new Set(days.map(day => day.work_date)).size === expectedDays;
  const allowed = new Set(scope.map(canonicalRestaurantName));
  return groups.filter(group => {
    if (scope.includes('*')) return true;
    const point = canonicalRestaurantName(group.name);
    return allowed.has(point) && points.get(point).size === 1;
  }).map(group => {
    const point = canonicalRestaurantName(group.name);
    const rows = points.get(point).size === 1 ? pointRows.get(point) || [] : [];
    const canCount = complete && rows.some(row => row.total > 0 || Object.values(row.details).some(items => items.length));
    const shifts = canCount ? rows.reduce((sum, row) => sum + row.total, 0) : null;
    const firstHalf = canCount ? rows.filter(row => Number(row.dateIso.slice(-2)) <= 15).reduce((sum, row) => sum + row.total, 0) : null;
    const salesRows = rows.filter(row => Number.isFinite(row.revenue));
    const revenue = salesRows.length === expectedDays ? salesRows.reduce((sum, row) => sum + row.revenue, 0) : null;
    return {
      name:group.name, accrual:group.accrual, shifts, firstHalf, secondHalf:canCount ? shifts - firstHalf : null, revenue, revenueDays:salesRows.length,
      unallocated:group.accrual - group.employees.reduce((sum, item) => sum + item.accrual, 0),
      employees:group.employees.map(({id, name, accrual, link}) => {
        const key = nameKey(name);
        const records = complete && rows.length && links.has(key) ? employeeRows.get(key) || [] : [];
        const local = records.filter(item => item.point === point);
        const firstHalf = local.length ? local.filter(item => item.day <= 15).reduce((sum, item) => sum + item.shift, 0) : null;
        const secondHalf = local.length ? local.filter(item => item.day > 15).reduce((sum, item) => sum + item.shift, 0) : null;
        return {id, name, accrual, link, firstHalf, secondHalf, comparable:local.length > 0 && records.length === local.length};
      }),
    };
  });
}
