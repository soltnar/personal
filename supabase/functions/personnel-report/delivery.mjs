// Same SaleOrder.List/Reglament=2 protocol as «История заказа».
const record=fields=>({_type:'record',f:0,s:fields.map(([t,n])=>({t,n})),d:fields.map(([, ,v])=>v)});
const unpack=value=>value?.s?Object.fromEntries(value.s.map((field,i)=>[field.n,value.d[i]])):value;
const rowsOf=result=>(result?.d||[]).map(d=>Object.fromEntries(result.s.map((field,i)=>[field.n,d[i]])));
export async function readDeliveryMetrics(from,to,env,request=fetch){
  const signal=AbortSignal.timeout(40000);
  const authResponse=await request('https://online.sbis.ru/oauth/service/',{method:'POST',signal,headers:{'Content-Type':'application/json'},body:JSON.stringify({app_client_id:env('SABY_APP_CLIENT_ID'),app_secret:env('SABY_APP_SECRET'),secret_key:env('SABY_SERVICE_KEY')})});
  const auth=await authResponse.json();
  if(!authResponse.ok||!auth.token)throw new Error('Не удалось авторизовать загрузку заказов доставки');
  const end=new Date(Date.parse(to)+86400000).toISOString().slice(0,10);
  let position=record([['Строка','NextDateWTZText',`${end} 00:00:00.000000`]]);
  const seen=new Set(),orders=[],positions=new Set([JSON.stringify(position)]);
  for(let page=0;page<200;page++){
    const response=await request('https://rest.saby.ru/service/',{method:'POST',signal,headers:{'Content-Type':'application/json; charset=UTF-8','X-SBISAccessToken':auth.token,'X-CalledMethod':'SaleOrder.List','X-Requested-With':'XMLHttpRequest',Referer:'https://rest.saby.ru/page/delivery'},body:JSON.stringify({jsonrpc:'2.0',protocol:7,id:page+1,method:'SaleOrder.List',params:{
      'Фильтр':record([['Строка','DateTimeStartWTZ',`${from} 00:00:00`],['Строка','DateTimeEndWTZ',`${to} 23:59:59`],['Строка','ProductStateId','Done'],['Число целое','Reglament',2]]),
      'Сортировка':{_type:'recordset',f:0,s:[{t:'Логическое',n:'l'},{t:'Строка',n:'n'},{t:'Логическое',n:'o'}],d:[[false,'NextDateWTZText',true]]},
      'Навигация':record([['Строка','Direction',page?'forward':'bothways'],['Логическое','HasMore',true],['Число целое','Limit',25],['Запись','Position',position]]),'ДопПоля':['RelatedSales','Reglament']}})});
    const data=await response.json();
    if(!response.ok||data.error||!Array.isArray(data.result?.s)||!Array.isArray(data.result?.d))throw new Error('Не удалось прочитать заказы доставки');
    const batch=rowsOf(data.result);
    const nav=Object.values(data.result).map(unpack).find(value=>value&&typeof value==='object'&&Object.hasOwn(value,'Position')&&Object.hasOwn(value,'HasMore'));
    if(!batch.length){if(nav?.HasMore===true)throw new Error('Пустая промежуточная страница доставки');break;}
    let added=0;
    for(const order of batch){if(order.Sale==null)throw new Error('Нет идентификатора заказа');const key=String(order.Sale);if(!seen.has(key)){seen.add(key);orders.push(order);added++;}}
    if(nav?.HasMore===false||data.result.n===false)break;
    if(!added)throw new Error('Saby повторяет страницу заказов');
    if(!nav&&batch.length<25)break;
    const next=nav?.Position|| (batch.at(-1).NextDateWTZText?record([['Строка','NextDateWTZText',batch.at(-1).NextDateWTZText]]):null);
    if(!next||positions.has(JSON.stringify(next)))throw new Error('Не продвинулся курсор заказов');
    positions.add(JSON.stringify(next));position=next;
    if(page===199)throw new Error('Превышен лимит страниц доставки');
  }
  const metrics=new Map();
  for(const order of orders){
    // Assigned courier is required: pickup and hall orders never enter this metric.
    if(!String(order.CourierName||'').trim())continue;
    const dateIso=String(order.DateWTZ||'').slice(0,10),restaurant=String(order.WarehouseName||order.Restaurant||order.SalesPointName||'').trim();
    if(!/^\d{4}-\d{2}-\d{2}$/.test(dateIso)||dateIso<from||dateIso>to||!restaurant)throw new Error('Не определены дата или ресторан курьерского заказа');
    const key=dateIso+'|'+restaurant;
    if(!metrics.has(key))metrics.set(key,{dateIso,restaurant,orders:0,revenue:0});
    const metric=metrics.get(key);metric.orders++;
    const amount=order.TotalPrice==null||order.TotalPrice===''?NaN:Number(String(order.TotalPrice).replace(/\s/g,'').replace(',','.'));
    if(!Number.isFinite(amount))metric.revenue=null;
    else if(metric.revenue!==null)metric.revenue+=amount;
  }
  return [...metrics.values()];
}
