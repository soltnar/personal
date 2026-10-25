import assert from 'node:assert/strict';
import {summarizeAttendance,restrictRows} from '../supabase/functions/personnel-report/engine.mjs';
const date='2026-10-23';
const marks=(department,position)=>[{department,position,person:department,dateTime:date+' 09:00:00',actionType:1},{department,position,person:department,dateTime:date+' 19:00:00',actionType:0}];
const days=[{work_date:date,rows:[...marks('Коминтерна 166 CALL CENTRE','администратор-кассир'),...marks('Самурай Белинского 61','Официант')]}];
const sales=[{sale_date:date,restaurant:'Самурай Белинского 61',revenue:445621},{sale_date:date,restaurant:'Коминтерна 166 CALL CENTRE',revenue:361916}];
const metrics=[{dateIso:date,restaurant:'Самурай Белинского 61 — Доставка',orders:129,revenue:253726.44}];
const network=[{dateIso:date,orders:1400,revenue:2600000}];
const rows=summarizeAttendance(days,sales,metrics,network,[date]);
const hall=rows.find(row=>row.restaurant==='Белинского, 61 Самурай'),cc=rows.find(row=>row.restaurant==='Коминтерна 166 CALL CENTRE');
assert.equal(hall.hallRevenue,191894.56);assert.equal(hall.revenue,445621);
assert.equal(cc.hall,0);assert.equal(cc.operators,1);assert.equal(cc.operatorsRevenue,2600000);assert.equal(cc.operatorsOrders,1400);assert.equal(cc.revenue,361916);assert.equal(cc.total,1);
const grant=restrictRows(rows,[{restaurant:cc.restaurant,roles:['hall']}],false)[0];assert.deepEqual(grant.allowedRoles,['operators']);assert.equal(grant.operatorsRevenue,2600000);assert(!Object.hasOwn(grant,'revenue'));assert(!Object.hasOwn(grant,'hall'));
assert.equal(restrictRows(rows,[{restaurant:hall.restaurant,roles:['hall']}],false)[0].hallRevenue,191894.56);
const missing=summarizeAttendance(days,sales).find(row=>row.restaurant===hall.restaurant);assert.equal(missing.hallRevenue,null,'Never use gross revenue when delivery is unknown');
console.log('PASS: hall excludes delivery, call-centre staff separated, full-network totals independent of gross revenue, legacy access preserved, missing source excluded');

const excluded=summarizeAttendance([{work_date:date,rows:marks('Самурай Ленина 36','логист').map(event=>({...event,person:'Кислицын Сергей Александрович'}))}],[])[0];
assert.equal(excluded.delivery,0);assert.equal(excluded.total,0);assert.equal(excluded.details.delivery.length,0);
console.log('PASS: owner-excluded Gazelle driver does not enter delivery staffing or productivity');

assert.equal(excluded.unclassified.length,1);
assert.equal(excluded.unclassified[0].excludedFromDelivery,true);
assert.equal(excluded.unclassified[0].marks.length,2);
assert.equal(excluded.unclassified[0].marks[0].location,'Ленина, 36 Самурай');
assert.match(excluded.unclassified[0].reviewReason,/Водитель газели/);
assert.equal(restrictRows([excluded],[],true)[0].unclassified.length,1);
assert(!Object.hasOwn(restrictRows([excluded],[{restaurant:excluded.restaurant,roles:['delivery']}],false)[0],'unclassified'));
console.log('PASS: excluded driver retains locations and entrance/exit marks in owner review without adding delivery shifts');

const technician=summarizeAttendance([{work_date:date,rows:marks('Самурай Доставка Моторный пер 2к1','Менеджер ремонтно-технического отд. упр. (Категория 3500/9)')}],[])[0];
assert.equal(technician.hall,0);assert.equal(technician.unclassified.length,1);
console.log('PASS: maintenance manager remains visible in review and never enters hall staffing');
