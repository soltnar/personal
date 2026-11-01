import assert from 'node:assert/strict';import fs from 'node:fs';import vm from 'node:vm';
import {summarizeAttendance} from '../supabase/functions/personnel-report/engine.mjs';
const date='2026-10-28',venue='Белинского, 61 Самурай';
const mark=(person,action,time)=>({person,actionType:action,dateTime:date+' '+time,position:'Повар',department:venue});
const events=[mark('Только вход',1,'09:00:00'),mark('Только выход',0,'20:00:00'),mark('Короткая смена',1,'09:00:00'),mark('Короткая смена',0,'12:00:00'),mark('Полная смена',1,'09:00:00'),mark('Полная смена',0,'20:00:00')];
const [row]=summarizeAttendance([{work_date:date,status:'ready',rows:events}],[]);
assert.equal(row.kitchen,3.5);assert.equal(row.hasConflict,true);
for(const [name,reason] of [['Только вход','нет отметки выхода'],['Только выход','нет отметки входа']]){const person=row.details.kitchen.find(p=>p.person===name);assert.equal(person.shift,1);assert.equal(person.hasConflict,true);assert.equal(person.reviewReason,reason);}
assert.equal(row.details.kitchen.find(p=>p.person==='Короткая смена').shift,.5);assert.equal(row.details.kitchen.find(p=>p.person==='Полная смена').hasConflict,false);
const source=fs.readFileSync(new URL('../script.js',import.meta.url),'utf8');
const context=vm.createContext({reportRows:null,dateSelect:'dates',restaurantSelect:'restaurants',getSelectedValues:key=>key==='dates'?[date]:[venue],getCheckedGroups:()=>['Кухня'],getRevenueFor:()=>100000});
vm.runInContext(source.slice(source.indexOf('function calcWorkedSeconds('),source.indexOf('function renderPeopleList(')),context);
const fallback=context.calculate(events.map(event=>({person:event.person,group:'Кухня',dateIso:date,restaurant:venue,direction:event.actionType===1?'Вход':'Выход',absSec:Date.parse(event.dateTime.replace(' ','T')+'+03:00')/1000})))[0];
assert.equal(fallback.kitchen,row.kitchen);assert(fallback.details.kitchen.filter(p=>p.person.startsWith('Только')).every(p=>p.shift===1&&p.hasConflict&&p.reviewReason));
console.log('PASS: a single entry or exit counts as one full shift with a warning in server and file-import calculations; normal full/half shifts unchanged');

for (const [name,times,expected] of [
 ['Повторный вход',[[1,'08:17:16'],[1,'23:01:52'],[0,'23:04:43']],1],
 ['Два интервала',[[1,'09:00:00'],[0,'12:00:00'],[1,'18:00:00'],[0,'20:00:00']],.5],
 ['Короткая с повтором',[[1,'16:00:00'],[1,'17:00:00'],[0,'20:00:00']],.5],
 ['Ночная с повтором',[[1,'16:00:00'],[1,'23:00:00'],[0,'2026-10-29 01:00:00']],1]
]) {
 const marks=times.map(([action,time])=>({...mark(name,action,time),dateTime:time.includes(' ')?time:date+' '+time}));
 const actual=summarizeAttendance([{work_date:date,status:'ready',rows:marks}],[])[0];
 assert.equal(actual.details.kitchen[0].shift,expected,name);
 const converted=marks.map(e=>({direction:e.actionType===1?'Вход':'Выход',absSec:Date.parse(e.dateTime.replace(' ','T')+'+03:00')/1000}));
 assert.equal(context.workedSecondsToShift(context.calcWorkedSeconds(converted)),expected,name+' import');
}
console.log('PASS: repeated entry preserves first entry; separate intervals exclude breaks; short and overnight shifts');
