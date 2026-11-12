/**
 * 웹 메뉴 수집기 → 구글 시트 진단표 자동 입력
 *
 * 설치 (진단표 구글 시트에서 1회)
 *   1. 확장 프로그램 → Apps Script → 이 파일 내용을 전부 붙여넣고 저장합니다.
 *   2. 위쪽 함수 목록에서 setup 을 선택하고 실행합니다. (권한 승인 → 실행 로그에 '연결 키'가 표시됩니다)
 *   3. 배포 → 새 배포 → 유형: 웹 앱 / 실행 사용자: 나 / 액세스 권한: 모든 사용자 → 배포
 *   4. 표시된 웹 앱 URL과 연결 키를 수집기 run.bat 의 '구글 시트 연결'에 입력합니다.
 *
 * 하는 일
 *   - 진단 시트(01_, 02_ …) 목록과 진단 대상·대표 도메인·기존 메뉴 수를 알려 줍니다.
 *   - 메뉴명을 13~15행(1·2·3 Depth)의 '마지막으로 사용한 열 다음 칸'부터 값으로만 입력합니다.
 *   - 메뉴·결과(20~67행)가 있는 열과 서식·수식·다른 칸은 건드리지 않습니다. 이미 있는 메뉴 경로는 건너뜁니다.
 */

var FIRST_COL = 8;      // H
var LAST_COL = 257;     // IW
var MENU_ROW = 13;      // 13~15행
var RESULT_ROW = 20;    // 20~67행
var RESULT_COUNT = 48;
var SITE_SHEET = /^\d\d_/;

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
  } catch (err) {
    return reply({ ok: false, error: '요청 형식이 올바르지 않습니다.' });
  }
  var token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || req.token !== token) {
    return reply({ ok: false, error: '연결 키가 맞지 않습니다. Apps Script에서 setup 을 실행해 연결 키를 확인하세요.' });
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (req.action === 'ping') return reply({ ok: true, title: ss.getName() });
    if (req.action === 'targets') return reply({ ok: true, title: ss.getName(), targets: listTargets(ss) });
    if (req.action === 'fill') return reply(fillSheet(ss, req));
    return reply({ ok: false, error: '알 수 없는 요청입니다: ' + req.action });
  } catch (err) {
    return reply({ ok: false, error: String((err && err.message) || err) });
  }
}

function doGet() {
  return reply({ ok: true, message: '웹 메뉴 수집기 연결용 주소입니다. 수집기에서 사용합니다.' });
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function listTargets(ss) {
  var out = [];
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (!SITE_SHEET.test(name) || name === '00_Sample') return;
    var head = sh.getRange('D3:D6').getDisplayValues();
    var area = readArea(sh);
    var info = planFill(area.menus, area.results, []);
    out.push({
      sheet: name,
      name: cleanText(head[0][0]) || name.replace(/^\d\d_/, ''),
      domain: cleanText(head[3][0]),
      menus: info.existing,
      start: info.start
    });
  });
  return out;
}

function fillSheet(ss, req) {
  var sh = ss.getSheetByName(req.sheet);
  if (!sh) return { ok: false, error: '시트를 찾을 수 없습니다: ' + req.sheet };
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return { ok: false, error: '다른 사용자가 입력 중입니다. 잠시 후 다시 시도하세요.' };
  try {
    var area = readArea(sh);
    var p = planFill(area.menus, area.results, req.menus || []);
    p.ok = true;
    p.sheet = req.sheet;
    p.range = p.write.length ? colName(p.start) + MENU_ROW + ':' + colName(p.start + p.write.length - 1) + (MENU_ROW + 2) : '';
    if (!req.dryRun && p.write.length) {
      var rows = [[], [], []];
      p.write.forEach(function (m) {
        for (var i = 0; i < 3; i++) rows[i].push(m[i] ? "'" + m[i] : '');   // ' = 글자로 입력 (수식·날짜로 바뀌지 않음)
      });
      sh.getRange(MENU_ROW, p.start, 3, p.write.length).setValues(rows);   // 값만 입력, 서식은 그대로
      var d6 = sh.getRange('D6');
      if (req.domain && cleanText(d6.getValue()) === '') {                 // 비어 있을 때만 대표 도메인 기록
        d6.setValue("'" + cleanText(req.domain));
        p.domainSaved = true;
      }
      SpreadsheetApp.flush();
    }
    p.written = req.dryRun ? 0 : p.write.length;
    return p;
  } finally {
    lock.releaseLock();
  }
}

function readArea(sh) {   // one read per sheet: rows 13~67, columns H~IW
  var rows = sh.getRange(MENU_ROW, FIRST_COL, RESULT_ROW + RESULT_COUNT - MENU_ROW, LAST_COL - FIRST_COL + 1).getValues();
  return { menus: rows.slice(0, 3), results: rows.slice(RESULT_ROW - MENU_ROW) };
}

/** Pure logic (tested with node): where the new menus go. menus/results are rows x columns from H. */
function planFill(menuRows, resultRows, menus) {
  var width = LAST_COL - FIRST_COL + 1;
  var used = -1;
  var existing = {};
  var count = 0;
  for (var c = 0; c < width; c++) {
    var key = [0, 1, 2].map(function (r) { return cleanText(menuRows[r][c]); });
    var hasMenu = key.some(function (x) { return x; });
    var hasResult = resultRows.some(function (row) { return cleanText(row[c]) !== ''; });
    if (hasMenu || hasResult) used = c;
    if (hasMenu && !existing[key.join('\u0001')]) {
      existing[key.join('\u0001')] = true;
      count++;
    }
  }
  var start = FIRST_COL + used + 1;
  var room = Math.max(0, LAST_COL - start + 1);
  var write = [], duplicates = [], overflow = [], seen = {};
  (menus || []).forEach(function (m) {
    var d = [0, 1, 2].map(function (i) { return cleanText((m || [])[i]); });
    if (!d[0]) return;
    var k = d.join('\u0001');
    if (existing[k] || seen[k]) { duplicates.push(d); return; }
    seen[k] = true;
    if (write.length < room) write.push(d); else overflow.push(d);
  });
  return { start: start, cell: colName(start) + MENU_ROW, existing: count, write: write, duplicates: duplicates, overflow: overflow };
}

function cleanText(v) {
  return String(v === null || v === undefined ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 240);
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
  module.exports = { planFill: planFill, cleanText: cleanText, colName: colName, FIRST_COL: FIRST_COL, LAST_COL: LAST_COL };
}
