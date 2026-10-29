import { canonicalRestaurantName } from '../personnel-report/engine.mjs';
const API='https://viruchki.ilgp.ru/api/telegram-settings';
export async function telegramDirectorySnapshot(authorization, request=fetch) {
  const read=async suffix=>{
    const response=await request(API+suffix,{headers:{Authorization:authorization},signal:AbortSignal.timeout(15000)});
    if(!response.ok)throw new Error('Не удалось загрузить пользователей Telegram из приложения выручки');
    return response.json();
  };
  const [snapshot,catalog]=await Promise.all([read(''),read('?action=catalog')]);
  if(!Array.isArray(snapshot.users)||!Array.isArray(catalog.warehouses)||!Array.isArray(catalog.restaurants))throw new Error('Некорректный справочник Telegram приложения выручки');
  const users=snapshot.users.filter(user=>Number.isSafeInteger(user.id)&&user.id>0).map(user=>{
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
    return {id:String(user.id),name:[user.firstName,user.lastName].filter(Boolean).join(' ')||user.username||`Telegram ${user.id}`,username:user.username||'',status:user.status,admin:!!user.admin,revision:user.revision,
      accessLabel:access.mode==='all'?'Все продажи':`Рестораны: ${(access.restaurants||[]).join(', ')||'не выбраны'}; отдельных складов: ${(access.warehouseIds||[]).length}`,
      suggestedAssignments:[...scope].map(([restaurant,roles])=>({restaurant,roles:[...roles]}))};
  });
  return {users,newUserRevision:snapshot.newUserRevision};
}

export async function readTelegramDirectory(authorization,request=fetch){return (await telegramDirectorySnapshot(authorization,request)).users;}
export async function changeTelegramAccess(authorization,body,request=fetch) {
  const id=Number(body.id);
  if(!Number.isSafeInteger(id)||id<=0||!['approved','revoked','rejected'].includes(body.status)||typeof body.revision!=='string')throw Object.assign(new Error('Проверьте Telegram ID и действие'),{status:400});
  const response=await request(API,{headers:{Authorization:authorization},signal:AbortSignal.timeout(15000)});
  if(!response.ok)throw new Error('Не удалось проверить пользователя Telegram');
  const snapshot=await response.json(),user=snapshot.users?.find(item=>item.id===id);
  if(body.revision!==(user?.revision||snapshot.newUserRevision))throw Object.assign(new Error('Данные сотрудника изменились. Обновите список.'),{status:409});
  if(user?.admin)throw Object.assign(new Error('Доступ администратора изменить нельзя'),{status:400});
  // Preserve all revenue permissions. Authorizing a new bot user grants no sales
  // access and enables no revenue mailing. Revocation disables that user's mailing.
  const payload={action:'user',id,revision:body.revision,status:body.status,
    salesAccess:user?.salesAccess||{mode:'limited',restaurants:[],warehouseIds:[]},
    mailing:{...(user?.mailing||{enabled:false,kind:'all',excludedRestaurants:[]}),...(body.status!=='approved'?{enabled:false}:{})}};
  const saved=await request(API,{method:'POST',headers:{Authorization:authorization,'Content-Type':'application/json'},body:JSON.stringify(payload),signal:AbortSignal.timeout(15000)});
  if(!saved.ok){const error=await saved.json();throw Object.assign(new Error(error.error||'Не удалось изменить доступ к боту'),{status:saved.status});}
  return {ok:true};
}
