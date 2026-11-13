/**
 * 웹 메뉴 수집기 -> 구글 시트 진단표 자동 입력
 *
 * 설치 (처음 1회, 어떤 진단표든 이 배포 하나로 씁니다)
 *   1. https://script.google.com 에서 새 프로젝트를 만들고 이 파일 내용을 전부 붙여넣고 저장합니다.
 *      (진단표의 확장 프로그램 -> Apps Script에 넣어도 됩니다. 그때는 그 진단표가 기본값이 됩니다.)
 *   2. 함수 목록에서 setup 을 선택해 실행합니다. 권한을 승인하면 실행 로그에 '연결 키'가 나옵니다.
 *   3. 배포 -> 새 배포 -> 유형: 웹 앱 / 실행 사용자: 나 / 액세스 권한: 모든 사용자 -> 배포 -> 웹 앱 URL 복사
 *   4. 수집기에서 run.bat --connect-google 로 웹 앱 URL과 연결 키를 입력합니다.
 *   실행할 때마다 수집기가 진단표 구글 시트 주소를 묻습니다(Enter = 지난번 시트).
 *
 * 하는 일
 *   - 진단 시트(01_, 02_ ...) 목록과 진단 대상(D3)·대표 도메인(D6)·기존 메뉴 수를 알려 줍니다.
 *   - 메뉴명을 13~15행(1·2·3 Depth)의 '마지막으로 사용한 열 다음 칸'부터 값으로만 입력합니다.
 *   - 메뉴·결과(20~67행)·수식·병합 셀이 있는 열과 숨긴 열, 서식, 다른 칸은 건드리지 않습니다.
 *     이미 있는 메뉴 경로는 건너뜁니다. 대표 도메인(D6)은 비어 있고 요청이 있을 때만 적습니다.
 */

var FIRST_COL = 8;      // H
var LAST_COL = 257;     // IW
var MENU_ROW = 13;      // 13~15행
var RESULT_ROW = 20;    // 20~67행
var RESULT_COUNT = 48;
var SITE_SHEET = /^\d{2,3}_/;
var VERSION = 8;        // the collector warns when the deployed web app is older than its own copy of this file

function setup() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    props.setProperty('TOKEN', token);
  }
  Logger.log('연결 키: ' + token);
  Logger.log('이제 배포 → 새 배포 → 웹 앱(실행: 나, 액세스: 모든 사용자)으로 배포하고 URL을 복사합니다.');
}

function resetToken() {
  PropertiesService.getScriptProperties().deleteProperty('TOKEN');
  setup();
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
    if (!req || typeof req !== 'object') throw new Error('empty');
  } catch (err) {
    return reply({ ok: false, error: '요청 형식이 올바르지 않습니다.' });
  }
  var token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || req.token !== token) {
    return reply({ ok: false, error: '연결 키가 맞지 않습니다. Apps Script에서 setup 을 실행해 연결 키를 확인하세요.' });
  }
  try {
    var ss = openBook(req.spreadsheet);
    if (req.action === 'ping') return reply({ ok: true, title: ss.getName(), url: ss.getUrl() });
    if (req.action === 'targets') return reply({ ok: true, title: ss.getName(), targets: listTargets(ss) });
    if (req.action === 'fill') return reply(fillSheet(ss, req));
    if (req.action === 'columns') return reply(menuColumns(ss, req));
    if (req.action === 'clear') return reply(clearColumns(ss, req));
    return reply({ ok: false, error: '알 수 없는 요청입니다: ' + req.action });
  } catch (err) {
    return reply({ ok: false, error: String((err && err.message) || err) });
  }
}

function openBook(ref) {
  ref = String(ref || '').trim();
  if (ref) {
    var m = ref.match(/\/d\/([a-zA-Z0-9_-]{20,})/);
    return m ? SpreadsheetApp.openById(m[1]) : (/^[a-zA-Z0-9_-]{20,}$/.test(ref) ? SpreadsheetApp.openById(ref) : SpreadsheetApp.openByUrl(ref));
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();       // bound to a sheet: that sheet by default
  if (!active) throw new Error('진단표 구글 시트 주소가 필요합니다.');
  return active;
}

function doGet() {
  return reply({ ok: true, message: '웹 메뉴 수집기 연결용 주소입니다. 수집기에서 사용합니다.' });
}

function reply(obj) {
  obj.version = VERSION;
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function listTargets(ss) {
  var out = [];
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (!isSiteSheet(sh)) return;
    var head = sh.getRange('D3:D6').getDisplayValues();
    var area = readArea(sh);
    var info = planFill(area.menus, area.results, []);
    out.push({
      sheet: name,
      name: cleanText(head[0][0]) || name.replace(/^\d{2,3}_/, ''),
      domain: cleanText(head[3][0], 2000),
      menus: info.existing,
      start: info.start,
      layout: layoutOk(sh)
    });
  });
  return out;
}

function layoutOk(sh) {    // rows 13~15 must be the 1/2/3 Depth rows of the diagnosis form
  var labels = sh.getRange('A13:A15').getDisplayValues().map(function (r) { return cleanText(r[0]); });
  return [1, 2, 3].every(function (n, i) { return new RegExp('^\\s*' + n + '\\s*(depth|단계|뎁스)', 'i').test(labels[i]); });
}

function isSiteSheet(sh) {
  var name = sh.getName();
  return SITE_SHEET.test(name) && name.toLowerCase() !== '00_sample' && !sh.isSheetHidden();
}

function fillSheet(ss, req) {
  var sh = ss.getSheetByName(req.sheet);
  if (!sh) return { ok: false, error: '시트를 찾을 수 없습니다: ' + req.sheet };
  if (!isSiteSheet(sh)) return { ok: false, error: "'" + req.sheet + "' 은(는) 진단 시트(01_ 형식, 숨김·00_Sample 제외)가 아닙니다." };
  if (!Array.isArray(req.menus) || req.menus.some(function (m) { return !Array.isArray(m); })) {
    return { ok: false, error: '메뉴 형식이 올바르지 않습니다.' };
  }
  if (!layoutOk(sh)) {
    return { ok: false, error: "'" + req.sheet + "' 시트의 13~15행이 1·2·3 Depth 행이 아닙니다. 진단표 양식이 달라 입력하지 않았습니다." };
  }
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();   // standalone deployment: script lock
  if (!lock.tryLock(30000)) return { ok: false, error: '다른 사용자가 입력 중입니다. 잠시 후 다시 시도하세요.' };
  try {
    var area = readArea(sh);
    var hidden = {};
    for (var c = area.start, need = (req.menus || []).length; c <= area.lastCol && need > 0; c++) {
      if (sh.isColumnHiddenByUser(c)) hidden[c] = true; else need--;
    }
    var p = planFill(area.menus, area.results, req.menus || [], area.lastCol, hidden);
    p.ok = true;
    p.sheet = req.sheet;
    p.gid = sh.getSheetId();
    p.url = ss.getUrl().replace(/#.*$/, '') + '#gid=' + p.gid;
    p.range = p.columns.length ? colName(p.columns[0]) + MENU_ROW + ':' + colName(p.columns[p.columns.length - 1]) + (MENU_ROW + 2) : '';
    if (!req.dryRun && p.write.length) {
      // another deployment (another account) has its own lock: the planned cells must still be empty right now
      var busy = p.columns.filter(function (col) {
        var cells = sh.getRange(MENU_ROW, col, 3, 1), v = cells.getValues(), f = cells.getFormulas();
        return v.some(function (r) { return cleanText(r[0]) !== ''; }) || f.some(function (r) { return r[0]; });
      });
      if (busy.length) return { ok: false, error: '입력하려던 칸에 그사이 다른 값이 들어왔습니다(' + busy.map(colName).join(', ') + '). 다시 실행하세요.' };
      // contiguous runs of columns (a hidden column splits a run); values only, formats stay
      for (var i = 0; i < p.write.length;) {
        var j = i;
        while (j + 1 < p.write.length && p.columns[j + 1] === p.columns[j] + 1) j++;
        var rows = [[], [], []];
        for (var k = i; k <= j; k++) {
          for (var r = 0; r < 3; r++) rows[r].push(p.write[k][r] ? "'" + p.write[k][r] : '');   // ' = 글자로 입력
        }
        sh.getRange(MENU_ROW, p.columns[i], 3, j - i + 1).setValues(rows);
        i = j + 1;
      }
    }
    if (!req.dryRun && cleanText(req.domain)) {
      var d6 = sh.getRange('D6');
      var merged = d6.getMergedRanges ? d6.getMergedRanges() : [];
      var topLeft = !merged.length || (merged[0].getRow() === 6 && merged[0].getColumn() === 4);
      if (topLeft && cleanText(d6.getValue()) === '' && d6.getFormula() === '') {   // 비어 있을 때만 대표 도메인 기록
        d6.setValue("'" + cleanText(req.domain, 2000));
        p.domainSaved = true;
      }
    }
    if (!req.dryRun) SpreadsheetApp.flush();
    p.written = req.dryRun ? 0 : p.write.length;
    return p;
  } finally {
    lock.releaseLock();
  }
}

/** Menu columns of a sheet for the clean-up check: [{col, menu: [d1, d2, d3], results: n, formula: bool}]. */
function menuColumns(ss, req) {
  var sh = ss.getSheetByName(req.sheet);
  if (!sh) return { ok: false, error: '시트를 찾을 수 없습니다: ' + req.sheet };
  if (!layoutOk(sh)) return { ok: false, error: "'" + req.sheet + "' 시트의 13~15행이 1·2·3 Depth 행이 아닙니다." };
  var area = readArea(sh), out = [];
  for (var c = 0; c <= area.lastCol - FIRST_COL; c++) {
    var menu = [0, 1, 2].map(function (r) { return cleanText(area.menus[r][c]); });
    if (!menu.some(function (x) { return x; })) continue;
    var results = area.results.filter(function (row) { return cleanText(row[c]) !== ''; }).length;
    out.push({ col: FIRST_COL + c, menu: menu, results: results });
  }
  return { ok: true, sheet: req.sheet, columns: out };
}

/** Empty rows 13~15 of the given columns, only when they still hold exactly the expected names and no result. */
function clearColumns(ss, req) {
  var sh = ss.getSheetByName(req.sheet);
  if (!sh) return { ok: false, error: '시트를 찾을 수 없습니다: ' + req.sheet };
  if (!layoutOk(sh)) return { ok: false, error: "'" + req.sheet + "' 시트의 13~15행이 1·2·3 Depth 행이 아닙니다." };
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { ok: false, error: '다른 사용자가 입력 중입니다. 잠시 후 다시 시도하세요.' };
  try {
    var area = readArea(sh), cleared = [], skipped = [];
    (req.columns || []).forEach(function (item) {
      var c = Number(item.col) - FIRST_COL;
      if (!(c >= 0 && c <= area.lastCol - FIRST_COL)) { skipped.push(item.col); return; }
      var now = [0, 1, 2].map(function (r) { return cleanText(area.menus[r][c]); });
      var want = [0, 1, 2].map(function (i) { return cleanText((item.menu || [])[i]); });
      var range = sh.getRange(MENU_ROW, Number(item.col), 3, 1);
      var formula = range.getFormulas().some(function (r) { return r[0]; });
      if (range.getMergedRanges().length || want.some(function (x) { return x === '=formula' || x === '=merged'; })) {
        skipped.push(item.col); return;                                      // merged labels belong to other columns too
      }
      var results = area.results.some(function (row) { return cleanText(row[c]) !== ''; });
      if (formula || results || now.join('\u0001') !== want.join('\u0001')) { skipped.push(item.col); return; }
      sh.getRange(MENU_ROW, Number(item.col), 3, 1).setValues([[''], [''], ['']]);   // values only, formats stay
      cleared.push(item.col);
    });
    SpreadsheetApp.flush();
    return { ok: true, sheet: req.sheet, cleared: cleared, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

function readArea(sh) {   // one read per sheet: rows 13~67, columns H~IW (clamped to the sheet's size)
  var height = RESULT_ROW + RESULT_COUNT - MENU_ROW;
  var width = Math.min(LAST_COL, sh.getMaxColumns()) - FIRST_COL + 1;
  var rowsAvail = Math.min(height, sh.getMaxRows() - MENU_ROW + 1);
  var full = LAST_COL - FIRST_COL + 1;
  var grid = [];
  for (var r = 0; r < height; r++) grid.push(new Array(full).fill(''));
  if (width > 0 && rowsAvail > 0) {
    var range = sh.getRange(MENU_ROW, FIRST_COL, rowsAvail, width);
    var values = range.getValues(), formulas = range.getFormulas();
    for (var i = 0; i < rowsAvail; i++) {
      for (var j = 0; j < width; j++) {
        var cached = values[i][j];
        grid[i][j] = cleanText(cached) !== '' ? cached : (formulas[i][j] ? '=formula' : '');   // a formula cell is used
      }
    }
    // merged cells: a merged 1 Depth label belongs to every column it covers; an empty merge is occupied
    range.getMergedRanges().forEach(function (m) {
      var r0 = m.getRow() - MENU_ROW, c0 = m.getColumn() - FIRST_COL;
      var anchor = (grid[r0] || [])[c0];
      for (var rr = r0; rr <= m.getLastRow() - MENU_ROW; rr++) {
        for (var cc = c0; cc <= m.getLastColumn() - FIRST_COL; cc++) {
          if (rr >= 0 && rr < height && cc >= 0 && cc < full && !(rr === r0 && cc === c0)) grid[rr][cc] = anchor || '=merged';
        }
      }
      if (!anchor && r0 >= 0 && c0 >= 0) grid[r0][c0] = '=merged';
    });
  }
  var lastCol = FIRST_COL + Math.max(width, 0) - 1;
  return { menus: grid.slice(0, 3), results: grid.slice(RESULT_ROW - MENU_ROW), lastCol: lastCol,
           start: planFill(grid.slice(0, 3), grid.slice(RESULT_ROW - MENU_ROW), [], lastCol).start };
}

/** Pure logic (tested with node): where the new menus go. menus/results are rows x columns from H. */
function planFill(menuRows, resultRows, menus, lastCol, hidden) {
  var width = LAST_COL - FIRST_COL + 1;
  var limit = Math.min(LAST_COL, lastCol || LAST_COL);   // never past the sheet's last column (no new columns)
  hidden = hidden || {};
  var used = -1;
  var existing = {};
  var count = 0;
  var marker = function (x) { return x === '=formula' || x === '=merged' ? '' : x; };
  for (var c = 0; c < width; c++) {
    var raw = [0, 1, 2].map(function (r) { return cleanText(menuRows[r][c]); });
    var key = raw.map(marker);
    var hasMenu = raw.some(function (x) { return x; });
    var hasResult = resultRows.some(function (row) { return cleanText(row[c]) !== ''; });
    if (hasMenu || hasResult) used = c;
    if (key.some(function (x) { return x; }) && !existing[key.join('\u0001')]) {
      existing[key.join('\u0001')] = true;
      count++;
    }
  }
  var start = FIRST_COL + used + 1;
  var free = [];
  for (var col = start; col <= limit; col++) if (!hidden[col]) free.push(col);
  var write = [], columns = [], duplicates = [], overflow = [], seen = {};
  (menus || []).forEach(function (m) {
    var d = [0, 1, 2].map(function (i) { return cleanText((m || [])[i]); });
    if (!d[0]) return;
    var k = d.join('\u0001');
    if (existing[k] || seen[k]) { duplicates.push(d); return; }
    seen[k] = true;
    if (write.length < free.length) { columns.push(free[write.length]); write.push(d); } else overflow.push(d);
  });
  return { start: start, cell: colName(start) + MENU_ROW, existing: count, write: write, columns: columns,
           duplicates: duplicates, overflow: overflow };
}

function cleanText(v, limit) {
  var s = String(v === null || v === undefined ? '' : v).replace(/[\u0000-\u001f\u007f-\u009f\ufeff]/g, ' ').replace(/\s+/g, ' ').trim();
  return Array.from(s).slice(0, limit || 240).join('');                     // by characters, never half an emoji
}

function colName(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

if (typeof module !== 'undefined') {
  module.exports = { VERSION: VERSION, planFill: planFill, cleanText: cleanText, colName: colName, FIRST_COL: FIRST_COL, LAST_COL: LAST_COL,
                     doPost: doPost, setup: setup };
}
