// node tests/js/menu_import.test.js  — pure planning logic of apps_script/MenuImport.gs
const fs = require('fs'); const path = require('path'); const vm = require('vm'); const assert = require('assert');
const src = fs.readFileSync(path.join(__dirname, '..', '..', 'apps_script', 'MenuImport.gs'), 'utf8');
const sandbox = { module: { exports: {} } }; vm.createContext(sandbox); vm.runInContext(src, sandbox);
const { planFill, colName } = sandbox.module.exports;
const W = 250, blank = n => Array.from({ length: n }, () => Array(W).fill(''));

assert.strictEqual(colName(8), 'H'); assert.strictEqual(colName(257), 'IW'); assert.strictEqual(colName(35), 'AI');

// empty sheet -> H13
let p = planFill(blank(3), blank(48), [['조회', '계좌조회', ''], ['조회', '계좌조회', ''], ['', 'x', '']]);
assert.strictEqual(p.cell, 'H13'); assert.strictEqual(p.write.length, 1); assert.strictEqual(p.duplicates.length, 1);

// menus in H..I, result only in K -> start L, existing path skipped
const m = blank(3), r = blank(48);
m[0][0] = '조회'; m[1][0] = ' 계좌조회 '; m[0][1] = '이체'; r[5][3] = '양호';
p = planFill(m, r, [['조회', '계좌조회'], ['고객센터', '공지사항', '']]);
assert.strictEqual(p.cell, 'L13'); assert.strictEqual(p.existing, 2);
assert.strictEqual(JSON.stringify(p.write), JSON.stringify([['고객센터', '공지사항', '']])); assert.strictEqual(p.duplicates.length, 1);

// full sheet -> overflow, nothing written
const f = blank(3); f[0][W - 1] = '끝';
p = planFill(f, blank(48), [['새 메뉴']]);
assert.strictEqual(p.write.length, 0); assert.strictEqual(p.overflow.length, 1);
// narrow sheet (columns only up to J): never past the last column
p = planFill(blank(3), blank(48), [['a'], ['b'], ['c'], ['d']], 10);
assert.strictEqual(p.write.length, 3); assert.strictEqual(p.overflow.length, 1);
// a formula cell counts as used
const fm = blank(3); fm[0][2] = '=formula';
p = planFill(fm, blank(48), [['x']]);
assert.strictEqual(p.cell, 'K13');
console.log('ok');
