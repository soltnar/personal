import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const source=fs.readFileSync(new URL('../script.js',import.meta.url),'utf8');
const context=vm.createContext({});
vm.runInContext(source.slice(source.indexOf('function yesterdayMoscow('),source.indexOf('function setDefaultSabyPeriod(')),context);
for(const [now,expected] of [['2026-10-22T08:00:00Z','2026-10-21'],['2026-10-21T21:01:00Z','2026-10-21'],['2026-01-01T00:00:00Z','2025-12-31'],['2024-03-01T00:00:00Z','2024-02-29']])assert.equal(context.yesterdayMoscow(new Date(now)),expected);
console.log('PASS: yesterday in Moscow, midnight boundary, year boundary, leap day');
