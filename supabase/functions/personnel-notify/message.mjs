import { restrictRows } from '../personnel-report/engine.mjs';
const number=value=>new Intl.NumberFormat('ru-RU',{maximumFractionDigits:2}).format(value);
export function kitchenMessage(rows, restaurant, date) {
  const selected=restrictRows(rows,[{restaurant,roles:['kitchen']}],false).filter(row=>row.dateIso===date);
  if(selected.length!==1)throw new Error('Нет однозначных данных предприятия за день');
  const row=selected[0], revenue=row.kitchenRevenue;
  const lines=[`${row.restaurant} · ${date.split('-').reverse().join('.')}`,'',`${row.kitchenRevenueSource==='unified'?'Выручка единого склада':'Выручка кухни'}: ${Number.isFinite(revenue)?number(revenue)+' ₽':'нет данных'}`,`Смены кухни: ${number(row.kitchen)}`,`Выработка: ${Number.isFinite(revenue)&&row.kitchen>0?number(revenue/row.kitchen)+' ₽/смену':'не рассчитана'}`,''];
  if(!Number.isFinite(revenue))lines.push('⚠ Выручка кухни не загружена. Сводка неполная.','');
  const issues=row.details.kitchen.filter(person=>person.hasConflict);
  lines.push(issues.length?'Персонал кухни — требуют проверки:':'Персонал кухни: замечаний по отметкам нет.');
  for(const person of issues)lines.push(`• ${person.person} — ${person.reviewReason||'проверьте отметки'}.`);
  lines.push('','Отчёт: https://soltnar.github.io/personal/');
  const text=lines.join('\n');
  if(text.length>4000)throw new Error('Список замечаний превышает размер сообщения; откройте отчёт');
  return text;
}
export function moscowDay(now=new Date()) {
  const today=new Intl.DateTimeFormat('en-CA',{timeZone:'Europe/Moscow',year:'numeric',month:'2-digit',day:'2-digit'}).format(now);
  const hour=Number(new Intl.DateTimeFormat('en-GB',{timeZone:'Europe/Moscow',hour:'2-digit',hourCycle:'h23'}).format(now));
  return {today,hour,yesterday:new Date(Date.parse(today+'T12:00:00Z')-86400000).toISOString().slice(0,10)};
}
