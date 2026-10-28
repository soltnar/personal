import { canonicalRestaurantName } from '../personnel-report/engine.mjs';
const API='https://viruchki.ilgp.ru/api/telegram-settings';
export async function readTelegramDirectory(authorization, request=fetch) {
  const read=async suffix=>{
    const response=await request(API+suffix,{headers:{Authorization:authorization},signal:AbortSignal.timeout(15000)});
    if(!response.ok)throw new Error('Не удалось загрузить пользователей Telegram из приложения выручки');
    return response.json();
  };
  const [snapshot,catalog]=await Promise.all([read(''),read('?action=catalog')]);
  if(!Array.isArray(snapshot.users)||!Array.isArray(catalog.warehouses)||!Array.isArray(catalog.restaurants))throw new Error('Некорректный справочник Telegram приложения выручки');
  return snapshot.users.filter(user=>Number.isSafeInteger(user.id)&&user.id>0).map(user=>{
    const access=user.salesAccess||{},scope=new Map();
    const add=(restaurant,role)=>{
      const name=canonicalRestaurantName(restaurant);
      if(!name||name==='Фабрика кухни'||name==='Подразделение не определено')return;
      if(!scope.has(name))scope.set(name,new Set());
      scope.get(name).add(role);
    };
    // Revenue permissions suggest only kitchen/bar. They do not grant access to
    // staff in the hall, delivery or to network-wide operator metrics.
    for(const warehouse of catalog.warehouses) {
      const allowed=access.mode==='all'||access.restaurants?.includes(warehouse.restaurant)||access.warehouseIds?.map(String).includes(String(warehouse.id));
      if(!allowed)continue;
      if(warehouse.kind==='kitchen'||warehouse.kind==='bar')add(warehouse.restaurant,warehouse.kind);
      if(['single','other'].includes(warehouse.kind)&&['Кварц, Гагарина, 224А','Моторный, 2/1 доставка Самурай','Коминтерна 166 CALL CENTRE'].includes(canonicalRestaurantName(warehouse.restaurant)))add(warehouse.restaurant,'kitchen');
    }
    return {id:String(user.id),name:[user.firstName,user.lastName].filter(Boolean).join(' ')||user.username||`Telegram ${user.id}`,username:user.username||'',status:user.status,admin:!!user.admin,
      accessLabel:access.mode==='all'?'Все продажи':`Рестораны: ${(access.restaurants||[]).join(', ')||'не выбраны'}; отдельных складов: ${(access.warehouseIds||[]).length}`,
      suggestedAssignments:[...scope].map(([restaurant,roles])=>({restaurant,roles:[...roles]}))};
  });
}
