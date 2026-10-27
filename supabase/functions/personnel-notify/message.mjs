import { restrictRows, canonicalRestaurantName, roleKeys } from '../personnel-report/engine.mjs';
const number=value=>new Intl.NumberFormat('ru-RU',{maximumFractionDigits:2}).format(value);
export const sectionLabels={kitchen:'Кухня',hall:'Зал',bar:'Бар',delivery:'Доставка',operators:'Операторы'};
const revenueFields={kitchen:'kitchenRevenue',hall:'hallRevenue',bar:'barRevenue',delivery:'deliveryRevenue',operators:'operatorsRevenue'};
const revenueLabels={kitchen:'Выручка кухни',hall:'Выручка зала без доставки',bar:'Выручка бара',delivery:'Сумма заказов доставки',operators:'Сумма заказов сети с самовывозом'};
export function subscriptionAssignments(value) {
  if(!Array.isArray(value)||value.length<1||value.length>18)throw new Error('Выберите предприятия и разделы');
  const names=new Set();
  return value.map(item=>{
    const restaurant=canonicalRestaurantName(item?.restaurant),roles=item?.roles;
    if(!restaurant||restaurant==='Подразделение не определено'||restaurant==='Фабрика кухни'||restaurant.length>200||names.has(restaurant)||!Array.isArray(roles)||!roles.length||roles.some(role=>!roleKeys.includes(role)))throw new Error('Проверьте предприятия и разделы');
    const callCentre=restaurant==='Коминтерна 166 CALL CENTRE';
    if((roles.includes('operators')&&!callCentre)||(roles.includes('hall')&&callCentre))throw new Error('Операторы доступны для колл-центра; зал — для ресторанов');
    names.add(restaurant);return {restaurant,roles:roleKeys.filter(role=>roles.includes(role))};
  });
}
export function reportMessage(rows, restaurant, date, roles=['kitchen']) {
  const [assignment]=subscriptionAssignments([{restaurant,roles}]);
  const selected=restrictRows(rows,[assignment],false).filter(row=>row.dateIso===date);
  if(selected.length!==1)throw new Error('Нет однозначных данных предприятия за день');
  const row=selected[0];
  const lines=[`${row.restaurant} · ${date.split('-').reverse().join('.')}`];
  for(const role of assignment.roles) {
    const revenue=row[revenueFields[role]],label=role==='kitchen'&&row.kitchenRevenueSource==='unified'?'Выручка единого склада':revenueLabels[role];
    lines.push('',sectionLabels[role],`${label}: ${Number.isFinite(revenue)?number(revenue)+' ₽':'нет данных'}`,`Смены: ${number(row[role])}`,`Выработка: ${Number.isFinite(revenue)&&row[role]>0?number(revenue/row[role])+' ₽/смену':'не рассчитана'}`);
    if(['delivery','operators'].includes(role)) {
      const orders=row[role==='delivery'?'deliveryOrders':'operatorsOrders'];
      lines.push(`Заказы: ${Number.isFinite(orders)?number(orders):'нет данных'}`);
    }
    if(!Number.isFinite(revenue))lines.push('⚠ Данные по выручке неполные. Сводка неполная.');
    const issues=(row.details[role]||[]).filter(person=>person.hasConflict);
    lines.push('',issues.length?'Персонал — требуют проверки:':'Персонал: замечаний по отметкам нет.');
    for(const person of issues)lines.push(`• ${person.person} — ${person.reviewReason||'проверьте отметки'}.`);
  }
  lines.push('','Отчёт: https://soltnar.github.io/personal/');
  const text=lines.join('\n');
  if(text.length>4000)throw new Error('Список замечаний превышает размер сообщения; откройте отчёт');
  return text;
}
// Retained for backwards-compatible callers.
export const kitchenMessage=(rows,restaurant,date)=>reportMessage(rows,restaurant,date,['kitchen']);
export function moscowDay(now=new Date()) {
  const today=new Intl.DateTimeFormat('en-CA',{timeZone:'Europe/Moscow',year:'numeric',month:'2-digit',day:'2-digit'}).format(now);
  const hour=Number(new Intl.DateTimeFormat('en-GB',{timeZone:'Europe/Moscow',hour:'2-digit',hourCycle:'h23'}).format(now));
  return {today,hour,yesterday:new Date(Date.parse(today+'T12:00:00Z')-86400000).toISOString().slice(0,10)};
}
