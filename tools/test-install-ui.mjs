import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const source=fs.readFileSync(new URL('../report-ui.js',import.meta.url),'utf8');
const installSource=source.slice(source.indexOf('let installEvent;'),source.indexOf("if('serviceWorker' in navigator)"));
function setup({standalone=false,ios=false}={}){
 const nodes=new Map(),listeners={},media={matches:standalone,addEventListener(type,fn){this.change=fn;}};
 const node=id=>{if(!nodes.has(id))nodes.set(id,{});return nodes.get(id);};
 const context=vm.createContext({document:{getElementById:node},navigator:{standalone:ios},window:{matchMedia:()=>media,addEventListener:(type,fn)=>{listeners[type]=fn;}},alert:()=>{throw Error('Unexpected installation instructions');}});
 vm.runInContext(installSource,context);return {node,listeners,media};
}
for(const options of [{standalone:true},{ios:true}]){
 const {node}=setup(options);assert.equal(node('installApp').hidden,true);assert.equal(node('installStatus').hidden,false);assert.match(node('installHint').textContent,/используете установленное/);
 await node('installApp').onclick();
}
const {node,listeners,media}=setup();assert.equal(node('installApp').hidden,false);assert.equal(node('installStatus').hidden,true);
let prompts=0,prevented=false;
listeners.beforeinstallprompt({preventDefault(){prevented=true;},async prompt(){prompts++;}});
await node('installApp').onclick();assert.equal(prevented,true);assert.equal(prompts,1);assert.equal(node('installApp').hidden,false,'A prompt alone is not proof of completed installation');
listeners.appinstalled();assert.equal(node('installApp').hidden,true);assert.equal(node('installStatus').hidden,false);
const changing=setup();changing.media.matches=true;changing.media.change();assert.equal(changing.node('installApp').hidden,true);
changing.media.matches=false;changing.media.change();assert.equal(changing.node('installApp').hidden,false);
console.log('PASS: standalone/iOS detection, install completion, display-mode changes and ordinary browser prompt');
