import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
class Element {
  constructor(){this.hidden=false;this.textContent='';this.attrs={};this.dataset={};this.children=[];this.classList={toggle(){},contains(){return false;}};}
  setAttribute(k,v){this.attrs[k]=v;} removeAttribute(k){delete this.attrs[k];} hasAttribute(k){return k in this.attrs;}
  append(...els){this.children.push(...els);} replaceChildren(){this.children=[];} focus(){}
  addEventListener(){}
}
const nodes=new Map(), node=id=>{if(!nodes.has(id))nodes.set(id,new Element());return nodes.get(id);};
const profileNames=[new Element(),new Element()],avatars=[new Element()],captions=[new Element()],app=new Element();
const pages=['overview','enterprises','quality','settings','users','account'].map(page=>{const el=new Element();el.dataset.pages=page;return el;});
const adminPage=new Element();adminPage.dataset.pages='users';adminPage.setAttribute('data-admin-only','');pages.push(adminPage);
const nav=pages.slice(0,5).map(el=>{const link=new Element();link.dataset.view=el.dataset.pages;return link;});
const document={body:new Element(),getElementById:node,createElement:()=>new Element(),addEventListener(){},querySelector:selector=>selector==='.app'?app:null,querySelectorAll:selector=>({'[data-pages]':pages,'.sidebarNav [data-view]':nav,'[data-profile-name]':profileNames,'[data-profile-avatar]':avatars,'[data-profile-caption]':captions}[selector]||[])};
const context=vm.createContext({document,sabyFromInput:{value:'2026-10-14'},sabyToInput:{value:'2026-10-21'},window:{innerWidth:1280,addEventListener(){},scrollTo(){}},location:{hash:'#overview'},getSabySession:()=>context.saby,revenueDbSession:null,saby:null,reportRows:[],reportAdmin:false,lastResultRows:[],roleLabels:{kitchen:'Кухня'},getCheckedGroups:()=>['Кухня'],productivitySample:()=>({missing:0}),prettyDate:x=>x,clearScopedSession:()=>{context.cleared++;context.reportRows=[];context.reportAdmin=false;},cleared:0,setTimeout,requestAnimationFrame:fn=>fn(),tableBody:{querySelectorAll:()=>[]}});
vm.runInContext(fs.readFileSync(new URL('../workspace.js',import.meta.url),'utf8'),context);
assert.equal(node('accountRole').textContent,'Гость');assert.equal(node('usersNav').hidden,true);
context.revenueDbSession={user:{id:'owner',email:'soltnar@gmail.com',user_metadata:{full_name:'Владелец сети'}}};
vm.runInContext('syncWorkspaceProfile()',context);
assert.equal(profileNames[0].textContent,'Владелец сети');assert.equal(node('accountProvider').textContent,'Google');assert.equal(node('usersNav').hidden,false);
context.location.hash='#users';vm.runInContext('showWorkspace("users")',context);assert.equal(adminPage.hidden,false);
context.revenueDbSession={user:{id:'staff',email:'staff@example.org',user_metadata:{full_name:'Иван Петров'}}};
vm.runInContext('syncWorkspaceProfile()',context);assert.equal(node('workspaceTitle').textContent,'Мой профиль');assert.equal(adminPage.hidden,true);assert.equal(context.cleared,1);
context.saby={subject:'123:456',name:'Анна Saby'};vm.runInContext('syncWorkspaceProfile()',context);
assert.equal(node('accountProvider').textContent,'Saby');assert.equal(profileNames[1].textContent,'Анна Saby');assert.equal(node('usersNav').hidden,true);
context.reportRows=[{restaurant:'Кафе A',allowedRoles:['kitchen']}];vm.runInContext('syncWorkspaceProfile()',context);assert.match(node('accountScope').textContent,/Кафе A.*Кухня/);
context.saby=null;context.revenueDbSession=null;vm.runInContext('syncWorkspaceProfile()',context);assert.equal(node('accountRole').textContent,'Гость');assert.equal(context.reportRows.length,0);
const html=fs.readFileSync(new URL('../index.html',import.meta.url),'utf8'),ids=[...html.matchAll(/\bid="([^"]+)"/g)].map(m=>m[1]);
assert.equal(new Set(ids).size,ids.length,'No duplicate IDs after separating screens');
console.log('PASS: sidebar routing, guest/Google/Saby profile, owner-only management, scoped labels, session clearing, unique IDs');
