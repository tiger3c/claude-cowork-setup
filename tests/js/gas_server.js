// A small Google Apps Script stand-in that runs apps_script/MenuImport.gs unchanged behind an HTTP server.
// node tests/js/gas_server.js <fixture.json>   -> prints "PORT <n>"; POST / = doPost, GET /__state = spreadsheets
// Fixture: {"token": "...", "books": {"<id>": {"name": "...", "sheets": [{"name", "hidden", "maxCols", "maxRows",
//           "cells": {"D6": "...", "H13": "..."}, "formulas": {"H16": "=..."}, "merges": ["H13:I13"], "hiddenCols": ["N"]}]}}}
const fs = require('fs'); const path = require('path'); const vm = require('vm'); const http = require('http');

const fixture = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const col = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const a1 = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return [Number(m[2]), col(m[1])]; };
const name = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

class Sheet {
  constructor(d) {
    Object.assign(this, { name: d.name, hidden: !!d.hidden, maxCols: d.maxCols || 260, maxRows: d.maxRows || 1000 });
    this.cells = {}; this.formulas = {};
    for (const [k, v] of Object.entries(d.cells || {})) this.cells[k] = v;
    for (const [k, v] of Object.entries(d.formulas || {})) this.formulas[k] = v;
    this.merges = (d.merges || []).map(r => { const [x, y] = r.split(':'); return [...a1(x), ...a1(y || x)]; });
    this.hiddenCols = new Set((d.hiddenCols || []).map(col));
    this.writes = [];
  }
  getName() { return this.name; } isSheetHidden() { return this.hidden; }
  getMaxColumns() { return this.maxCols; } getMaxRows() { return this.maxRows; }
  isColumnHiddenByUser(c) { return this.hiddenCols.has(c); }
  getRange(r, c, nr, nc) {
    if (typeof r === 'string') { const [x, y] = r.split(':'); const [r1, c1] = a1(x); const [r2, c2] = a1(y || x); return new Range(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1); }
    if (c + (nc || 1) - 1 > this.maxCols || r + (nr || 1) - 1 > this.maxRows) throw new Error('범위가 시트 크기를 벗어났습니다');
    return new Range(this, r, c, nr || 1, nc || 1);
  }
}
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  each(fn) { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(fn(name(this.c + j) + (this.r + i))); out.push(row); } return out; }
  getValues() { return this.each(k => this.sh.cells[k] ?? ''); }
  getDisplayValues() { return this.each(k => String(this.sh.cells[k] ?? '')); }
  getFormulas() { return this.each(k => this.sh.formulas[k] || ''); }
  getValue() { return this.getValues()[0][0]; } getFormula() { return this.getFormulas()[0][0]; }
  setValue(v) { this.setValues([[v]]); }
  setValues(rows) {
    if (rows.length !== this.nr || rows.some(r => r.length !== this.nc)) throw new Error('setValues 크기가 범위와 다릅니다');
    rows.forEach((row, i) => row.forEach((v, j) => {
      const k = name(this.c + j) + (this.r + i);
      if (this.sh.formulas[k]) throw new Error('수식 칸에 값을 썼습니다: ' + k);
      const text = typeof v === 'string' && v.startsWith("'") ? v.slice(1) : v;   // Sheets: leading ' = plain text
      if (text === '') delete this.sh.cells[k]; else this.sh.cells[k] = text;
      this.sh.writes.push(k);
    }));
  }
  getMergedRanges() {
    return this.sh.merges.filter(([r1, c1, r2, c2]) => r1 <= this.r + this.nr - 1 && r2 >= this.r && c1 <= this.c + this.nc - 1 && c2 >= this.c)
      .map(([r1, c1, r2, c2]) => ({ getRow: () => r1, getColumn: () => c1, getLastRow: () => r2, getLastColumn: () => c2 }));
  }
}
const books = {};
for (const [id, b] of Object.entries(fixture.books)) {
  const sheets = b.sheets.map(s => new Sheet(s));
  books[id] = { getName: () => b.name, getUrl: () => `https://docs.google.com/spreadsheets/d/${id}/edit`, getSheets: () => sheets,
                getSheetByName: n => sheets.find(s => s.name === n) || null, sheets };
}
const props = { TOKEN: fixture.token };
const context = {
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; }, deleteProperty: k => { delete props[k]; } }) },
  Utilities: { getUuid: () => '00000000-0000-4000-8000-' + String(Date.now()).padStart(12, '0') },
  Logger: { log: () => {} },
  LockService: { getDocumentLock: () => null, getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  SpreadsheetApp: {
    openById: id => { if (!books[id]) throw new Error('문서를 찾을 수 없습니다: ' + id); return books[id]; },
    openByUrl: url => context.SpreadsheetApp.openById((/\/d\/([^/]+)/.exec(url) || [])[1]),
    getActiveSpreadsheet: () => null, flush: () => {},
  },
  ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ content: s, setMimeType() { return this; } }) },
  module: { exports: {} }, console,
};
vm.createContext(context);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', '..', 'apps_script', 'MenuImport.gs'), 'utf8'), context);
const { doPost } = context.module.exports;

http.createServer((req, res) => {
  if (req.method === 'GET' && req.url === '/__state') {
    const out = {};
    for (const [id, b] of Object.entries(books)) out[id] = b.sheets.map(s => ({ name: s.name, cells: s.cells, writes: s.writes }));
    res.setHeader('Content-Type', 'application/json'); return res.end(JSON.stringify(out));
  }
  let body = '';
  req.on('data', d => { body += d; });
  req.on('end', () => {
    const out = doPost({ postData: { contents: body } });
    res.setHeader('Content-Type', 'application/json'); res.end(out.content);
  });
}).listen(0, '127.0.0.1', function () { console.log('PORT ' + this.address().port); });
