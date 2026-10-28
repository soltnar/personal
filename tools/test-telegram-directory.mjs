import assert from 'node:assert/strict';
import {readTelegramDirectory} from '../supabase/functions/personnel-notify/directory.mjs';
const catalog={restaurants:['Самурай Белинского 61','Самурай Ленина 36'],warehouses:[
 {id:'k',name:'Кухня',restaurant:'Самурай Белинского 61',kind:'kitchen'},
 {id:'b',name:'Бар',restaurant:'Самурай Белинского 61',kind:'bar'},
 {id:'l',name:'Кухня',restaurant:'Самурай Ленина 36',kind:'kitchen'},
 {id:'c',name:'Единый',restaurant:'Коминтерна 166',kind:'other'}]};
const users=[{id:123456,firstName:'Александр',lastName:'Шалаев',status:'approved',salesAccess:{mode:'limited',restaurants:[],warehouseIds:['k']}},
 {id:234567,status:'approved',salesAccess:{mode:'limited',restaurants:['Самурай Ленина 36'],warehouseIds:[]}},
 {id:345678,status:'approved',salesAccess:{mode:'all'}},
 {id:456789,status:'revoked',salesAccess:{mode:'limited',restaurants:[],warehouseIds:[]}}];
let requests=[];
const result=await readTelegramDirectory('Bearer owner',async(url,init)=>{requests.push({url,authorization:init.headers.Authorization});return Response.json(url.includes('catalog')?catalog:{users});});
assert.equal(result[0].name,'Александр Шалаев');assert.equal(result[0].id,'123456');
assert.deepEqual(result[0].suggestedAssignments,[{restaurant:'Белинского, 61 Самурай',roles:['kitchen']}]);
assert.deepEqual(result[1].suggestedAssignments,[{restaurant:'Ленина, 36 Самурай',roles:['kitchen']}]);
assert(result[2].suggestedAssignments.some(row=>row.restaurant==='Коминтерна 166 CALL CENTRE'&&row.roles.includes('kitchen')));
assert(result.every(user=>user.suggestedAssignments.every(row=>row.roles.every(role=>['kitchen','bar'].includes(role)))));
assert.equal(result[3].status,'revoked');assert.equal(result[3].suggestedAssignments.length,0);
assert(requests.every(item=>item.authorization==='Bearer owner'));
await assert.rejects(readTelegramDirectory('Bearer owner',async()=>new Response('',{status:403})),/Не удалось загрузить/);
console.log('PASS: central Telegram directory, owner authorization forwarded, warehouse/restaurant permissions, unified kitchens, no inferred hall/operator personnel access, source failure');
