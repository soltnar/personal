import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
class Element{
 constructor(){this.children=[];this.value='period';this.classList={add(){}};}
 addEventListener(){} append(...children){this.children.push(...children);} replaceChildren(){this.children=[];}
}
const nodes=new Map(),node=id=>{if(!nodes.has(id))nodes.set(id,new Element());return nodes.get(id);};
const context=vm.createContext({SUPABASE_URL:'https://test.invalid',supabaseClient:null,revenueDbSession:null,
 document:{getElementById:node,querySelectorAll:()=>[],createElement:()=>new Element()},window:{addEventListener(){}},navigator:{},
 getCheckedGroups:()=>['Кухня','Зал','Бар','Доставка'],formatMoney:value=>String(value),formatShift:value=>String(value??'—'),prettyDate:value=>value,
 getSelectedValues:()=>['Фабрика кухни','Ресторан'],restaurantSelect:{},reportRows:[],sabyFromInput:{value:'2026-10-01'},sabyToInput:{value:'2026-10-02'}});
vm.runInContext(fs.readFileSync(new URL('../report-ui.js',import.meta.url),'utf8'),context);
const factory={restaurant:'Фабрика кухни',dateIso:'2026-10-01',kitchen:14,hall:0,bar:0,delivery:0,revenue:null,kitchenRevenue:null,barRevenue:null,allowedRoles:['kitchen'],details:{kitchen:[{person:'Сотрудник',shift:1}],hall:[],bar:[],delivery:[]}};
const restaurant={...factory,restaurant:'Ресторан',kitchen:4,revenue:100000,kitchenRevenue:80000};
assert.equal(context.productivitySample(factory,'kitchen').missing,0);
assert.equal(context.productivitySample({...factory,kitchenRevenue:999999},'kitchen').shifts,0,'Production never enters restaurant productivity');
assert.equal(context.productivitySample({...restaurant,kitchenRevenue:null},'kitchen').missing,1,'Actual restaurant gaps stay visible');
context.reportRows=[factory,restaurant];
const summary=context.summarizeReportRevenue([factory,restaurant]);assert.equal(summary.partial,false);assert.equal(summary.known.length,1);
assert.equal(context.summarizeReportRevenue([factory,{...restaurant,revenue:null}]).partial,true);
context.renderGaps([factory]);assert.match(node('reportGaps').children[0].textContent,/— 0 записей/);
context.renderProductivity([factory]);assert.equal(node('productivity').children.length,1);assert.match(node('productivity').children[0].textContent,/производство/);
context.renderProductivity([factory,restaurant]);assert.match(node('productivity').children[0].innerHTML,/20000 ₽/);assert(!node('productivity').children[0].innerHTML.includes('coverageWarning'));
const details=context.roleMetricsHtml(factory);assert.match(details,/<dt>Смены<\/dt><dd>14/);assert(!details.includes('<dt>Выручка'));assert(!details.includes('<dt>Выработка'));
const exported=context.reportExportData([factory])[1];assert.equal(exported[2],14);assert(exported.slice(6).every(value=>value===''),'Production export keeps shifts but leaves financial fields inapplicable');
const result=context.calculateScopedRows();const production=result.find(row=>row.restaurant==='Фабрика кухни');assert.equal(production.kitchen,14);assert.equal(production.total,14);assert.equal(production.coverage.kitchen.missing,0);assert.equal(production.coverage.kitchen.shifts,0);
console.log('PASS: production staffing preserved; no revenue gaps, financial warnings, productivity denominator or financial export values; restaurant metrics unchanged');
