"use strict";

const $ = sel => document.querySelector(sel);
const $$ = sel => [...document.querySelectorAll(sel)];

const token = new URLSearchParams(location.hash.slice(1)).get("token") || sessionStorage.getItem("menu-token") || "";
if (token) sessionStorage.setItem("menu-token", token);

const MODE_TITLES = {
  public: "로그인 전 메뉴 수집",
  login: "로그인 후 메뉴 수집",
  compare: "로그인 전·후 비교 수집",
};
const STATE_LABELS = {
  idle: "대기", starting: "브라우저 여는 중", ready: "준비됨", crawling: "수집 중",
  capturing: "화면 읽는 중", complete: "완료", stopped: "중지됨", error: "오류",
};
const BUSY = new Set(["starting", "crawling", "capturing"]);

const app = {
  mode: "public",
  report: null,
  timer: null,
  autoStarted: false,
  resultShown: false,
  edits: new Map(),    // id -> [d1, d2, d3]
  checked: new Map(),  // id -> bool
};

async function api(path, body) {
  const res = await fetch(path, {
    method: body === undefined ? "GET" : "POST",
    headers: { "x-session-token": token, ...(body === undefined ? {} : { "content-type": "application/json" }) },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (!res.ok) {
    let text = await res.text();
    try {
      const data = JSON.parse(text);
      text = typeof data.detail === "string" ? data.detail
        : Array.isArray(data.detail) ? data.detail.map(d => `${(d.loc || []).slice(-1)[0]}: ${d.msg}`).join(" / ") : text;
    } catch { /* plain text */ }
    throw new Error(text || `요청 실패 (${res.status})`);
  }
  return res;
}

function showError(message) {
  const box = $("#error");
  box.textContent = message || "";
  box.hidden = !message;
  if (message) window.scrollTo({ top: 0, behavior: "smooth" });
}

async function run(task) {
  showError("");
  try { await task(); } catch (err) { showError(err.message); }
}

function goStep(step) {
  $$("[data-panel]").forEach(p => { p.hidden = p.dataset.panel !== String(step); });
  $$("#steps li").forEach(li => {
    const n = Number(li.dataset.step);
    li.classList.toggle("active", n === step);
    li.classList.toggle("done", n < step);
  });
}

// ---------- 1·2단계: 방식 선택과 정보 입력 ----------

function applyMode() {
  app.mode = $('input[name="mode"]:checked').value;
  $$("[data-modes]").forEach(el => { el.hidden = !el.dataset.modes.split(" ").includes(app.mode); });
  $("#formTitle").textContent = `${MODE_TITLES[app.mode]} · 정보 입력`;
  $$(".mode").forEach(m => m.classList.toggle("selected", m.querySelector("input").checked));
}

function readConfig() {
  const f = $("#form").elements;
  const num = name => Number(f[name].value);
  return {
    mode: app.mode,
    url: f.url.value.trim(),
    site_name: f.site_name.value.trim() || "메뉴 수집",
    login_url: app.mode === "public" ? "" : f.login_url.value.trim(),
    browser: f.browser.value,
    headless: app.mode === "public" && f.headless.checked,
    extra_hosts: f.extra_hosts.value.split(/[\s,]+/).map(x => x.trim()).filter(Boolean),
    menu_selector: f.menu_selector.value.trim(),
    max_pages: num("max_pages"), max_seconds: num("max_seconds"), max_hops: num("max_hops"),
    wait_ms: num("wait_ms"), delay_ms: num("delay_ms"), timeout_ms: num("timeout_ms"),
  };
}

async function openBrowser(event) {
  event.preventDefault();
  const f = $("#form").elements;
  if (!f.url.value.trim()) { showError("사이트 주소를 입력하세요."); f.url.focus(); return; }
  if (!/^https?:\/\//i.test(f.url.value.trim())) { showError("사이트 주소는 http:// 또는 https://로 시작해야 합니다."); f.url.focus(); return; }
  if (!f.consent.checked) { showError("점검 권한 확인에 체크하세요."); return; }
  await run(async () => {
    $("#openBtn").disabled = true;
    try {
      await api("/api/open", readConfig());
    } finally {
      $("#openBtn").disabled = false;
    }
    app.autoStarted = false;
    app.resultShown = false;
    app.edits.clear();
    app.checked.clear();
    app.report = null;
    $("#runTitle").textContent = MODE_TITLES[app.mode];
    goStep(3);
    startPolling();
  });
}

// ---------- 3단계: 진행 ----------

function startPolling() {
  clearInterval(app.timer);
  poll();
  app.timer = setInterval(poll, 1000);
}

async function poll() {
  try {
    const report = await (await api("/api/status")).json();
    app.report = report;
    renderRun(report);
    autoFlow(report);
  } catch (err) {
    showError(err.message);
    clearInterval(app.timer);
  }
}

function phaseStage(report) {
  // What the user should do next in the current mode.
  const mode = report.config?.mode || app.mode;
  const done = report.phases_done || [];
  if (report.state === "error") return "error";
  if (BUSY.has(report.state)) return "busy";
  if (mode === "public") return done.includes("before") ? "finished" : "crawl-before";
  if (mode === "login") return done.includes("after") ? "finished" : "login";
  if (!done.includes("before")) return "crawl-before";
  return done.includes("after") ? "finished" : "login";
}

function renderTimeline(report) {
  const mode = report.config?.mode || app.mode;
  const done = report.phases_done || [];
  const steps = {
    public: [["before", "로그인 전 메뉴 수집"]],
    login: [["login", "브라우저에서 로그인"], ["after", "로그인 후 메뉴 수집"]],
    compare: [["before", "1. 로그인 전 수집"], ["login", "2. 로그인"], ["after", "3. 로그인 후 수집"]],
  }[mode];
  const current = BUSY.has(report.state) ? report.phase : phaseStage(report) === "login" ? "login" : "";
  $("#timeline").replaceChildren(...steps.map(([key, label]) => {
    const li = document.createElement("li");
    li.textContent = label;
    const isDone = key === "login" ? done.includes("after") || (mode === "login" && BUSY.has(report.state) && report.phase === "after")
      : done.includes(key) && !(BUSY.has(report.state) && report.phase === key);
    li.classList.toggle("done", isDone);
    li.classList.toggle("active", key === current);
    return li;
  }));
}

function renderRun(report) {
  const menus = report.menus || [];
  $("#stateBadge").textContent = STATE_LABELS[report.state] || report.state;
  $("#stateBadge").className = `badge ${report.state}`;
  $("#message").textContent = report.message || "";
  $("#statMenus").textContent = menus.length;
  $("#statReady").textContent = menus.filter(m => m.ready).length;
  const maxPages = report.config?.max_pages;
  $("#statPages").textContent = maxPages ? `${report.visited_count || 0} / ${maxPages}` : report.visited_count || 0;
  $("#statIssues").textContent = (report.issues || []).filter(x => !x.resolved).length;
  $("#current").textContent = report.current_url ? `확인 중: ${report.current_url}` : "";
  $("#loginCheck").textContent = report.login_check || "";
  $("#loginCheck").hidden = !report.login_check;
  renderTimeline(report);

  const stage = phaseStage(report);
  const busy = BUSY.has(report.state);
  $("#loginBox").hidden = stage !== "login";
  $("#stopBtn").hidden = report.state !== "crawling";
  $("#captureBtn").hidden = busy || stage === "error" || report.state === "idle";
  $("#toResult").hidden = !(report.phases_done || []).length || busy;
  $("#recrawlBtn").hidden = stage !== "finished";
  $("#loginDone").disabled = busy;
}

async function control(action, phase) {
  await api(`/api/control/${action}`, phase ? { phase } : {});
  poll();
}

function autoFlow(report) {
  const stage = phaseStage(report);
  if (stage === "crawl-before" && report.state === "ready" && !app.autoStarted) {
    app.autoStarted = true;
    run(() => control("crawl", "before"));
  }
  if (stage === "finished" && !app.resultShown && ["complete", "stopped"].includes(report.state)) {
    app.resultShown = true;
    showResult();
  }
}

function currentPhase() {
  const mode = app.report?.config?.mode || app.mode;
  if (mode === "public") return "before";
  if (mode === "login") return "after";
  return (app.report?.phases_done || []).includes("before") ? "after" : "before";
}

async function fillLogin(event) {
  event.preventDefault();
  const form = event.target;
  const username = form.elements.username.value;
  const password = form.elements.password.value;
  form.elements.password.value = "";
  if (!password) { showError("비밀번호를 입력하세요."); return; }
  await run(() => api("/api/login/fill", { username, password }));
  form.elements.username.value = "";
  poll();
}

// ---------- 4단계: 결과 ----------

function showResult() {
  goStep(4);
  renderResult();
}

function depthsOf(m) {
  return app.edits.get(m.id) || (m.path.concat(["", "", ""])).slice(0, 3);
}

function isChecked(m) {
  return app.checked.has(m.id) ? app.checked.get(m.id) : m.ready;
}

function visibleMenus() {
  const menus = app.report?.menus || [];
  const text = $("#fText").value.trim().toLowerCase();
  const status = $("#fStatus").value;
  const vis = $("#fVis").value;
  const onlyChecked = $("#fChecked").checked;
  return menus.filter((m, i) => {
    m._no = i + 1;
    if (status && m.status !== status) return false;
    if (vis && m.visibility !== vis) return false;
    if (onlyChecked && !isChecked(m)) return false;
    if (text && !(depthsOf(m).join(" ") + " " + m.url).toLowerCase().includes(text)) return false;
    return true;
  });
}

function cell(text, cls) {
  const td = document.createElement("td");
  td.textContent = text ?? "";
  if (cls) td.className = cls;
  return td;
}

function renderResult() {
  const report = app.report;
  if (!report) return;
  const compare = report.config?.mode === "compare";
  $("#fVis").hidden = !compare;
  document.body.classList.toggle("compare", compare);

  const statuses = [...new Set(report.menus.map(m => m.status))].sort();
  const sel = $("#fStatus");
  const keep = sel.value;
  sel.replaceChildren(new Option("모든 상태", ""), ...statuses.map(s => new Option(s, s)));
  sel.value = statuses.includes(keep) ? keep : "";

  const rows = visibleMenus().slice(0, 3000).map(m => {
    const tr = document.createElement("tr");
    tr.className = m.ready ? "ok" : m.kind === "group" ? "group" : "warn";
    const box = document.createElement("input");
    box.type = "checkbox";
    box.checked = isChecked(m);
    box.addEventListener("change", () => { app.checked.set(m.id, box.checked); updateCount(); });
    const tdBox = document.createElement("td");
    tdBox.append(box);
    tr.append(tdBox, cell(m._no));
    const depths = depthsOf(m);
    depths.forEach((value, idx) => {
      const input = document.createElement("input");
      input.value = value;
      input.maxLength = 240;
      input.addEventListener("input", () => {
        const next = depthsOf(m).slice();
        next[idx] = input.value;
        app.edits.set(m.id, next);
      });
      const td = document.createElement("td");
      td.append(input);
      tr.append(td);
    });
    tr.append(cell(m.status, "status"), cell(m.visibility, "vis-col"), cell(m.evidence));
    const urlTd = cell(m.url, "url");
    urlTd.title = m.reason || m.url;
    tr.append(urlTd);
    return tr;
  });
  if (!rows.length) {
    const tr = document.createElement("tr");
    const td = cell("조건에 맞는 메뉴가 없습니다.", "empty");
    td.colSpan = 9;
    tr.append(td);
    rows.push(tr);
  }
  $("#rows").replaceChildren(...rows);

  const issues = (report.issues || []).map(x => {
    const tr = document.createElement("tr");
    if (x.resolved) tr.className = "ok";
    tr.append(cell(x.phase), cell(x.kind), cell(x.url, "url"), cell(x.reason), cell(x.resolved));
    return tr;
  });
  if (!issues.length) {
    const tr = document.createElement("tr");
    const td = cell("페이지 단위 확인 사항이 없습니다.", "empty");
    td.colSpan = 5;
    tr.append(td);
    issues.push(tr);
  }
  $("#issues").replaceChildren(...issues);
  updateCount();
}

function updateCount() {
  const menus = app.report?.menus || [];
  const n = menus.filter(isChecked).length;
  $("#selCount").textContent = `체크 ${n}개 / 전체 ${menus.length}개 · 진단 시트 ${Math.max(1, Math.ceil(n / 250))}묶음`;
}

function setVisibleChecked(value) {
  visibleMenus().forEach(m => app.checked.set(m.id, value));
  renderResult();
}

async function exportExcel() {
  await run(async () => {
    const selections = (app.report?.menus || []).filter(isChecked).map(m => ({ id: m.id, depths: depthsOf(m).map(x => x.trim()) }));
    const empty = selections.findIndex(s => !s.depths[0]);
    if (empty >= 0) throw new Error("체크한 메뉴 중 1 Depth가 비어 있는 항목이 있습니다. 메뉴명을 입력하거나 체크를 해제하세요.");
    const res = await api("/api/export", { selections });
    const blob = await res.blob();
    const name = (app.report.config.site_name || "메뉴수집").replace(/[\\/:*?"<>|]+/g, "_");
    const stamp = new Date().toISOString().slice(0, 16).replace(/[-:T]/g, "");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${name}_메뉴목록_${stamp}.xlsx`;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  });
}

// ---------- 연결 ----------

document.addEventListener("DOMContentLoaded", async () => {
  if (!token) {
    showError("실행 링크의 접속 정보가 없습니다. ‘run.bat’ 창에 표시된 주소로 다시 접속하세요.");
    return;
  }
  $$('input[name="mode"]').forEach(r => r.addEventListener("change", applyMode));
  applyMode();
  $("#toForm").addEventListener("click", () => { applyMode(); goStep(2); $("#form").elements.url.focus(); });
  $("#backToMode").addEventListener("click", () => goStep(1));
  $("#form").addEventListener("submit", openBrowser);
  $("#loginOpen").addEventListener("click", () => run(() => control("login")));
  $("#loginDone").addEventListener("click", () => run(() => control("crawl", "after")));
  $("#fillForm").addEventListener("submit", fillLogin);
  $("#captureBtn").addEventListener("click", () => run(() => control("capture", currentPhase())));
  $("#recrawlBtn").addEventListener("click", () => run(() => { app.resultShown = false; return control("crawl", currentPhase()); }));
  $("#stopBtn").addEventListener("click", () => run(() => control("stop")));
  $("#toResult").addEventListener("click", showResult);
  $("#backToRun").addEventListener("click", () => goStep(3));
  $("#exportBtn").addEventListener("click", exportExcel);
  ["#fText", "#fStatus", "#fVis", "#fChecked"].forEach(s => $(s).addEventListener("input", renderResult));
  $("#checkAll").addEventListener("click", () => setVisibleChecked(true));
  $("#uncheckAll").addEventListener("click", () => setVisibleChecked(false));
  $("#restart").addEventListener("click", () => run(async () => {
    if (!confirm("수집 결과와 브라우저를 닫고 처음으로 돌아갈까요? 엑셀을 먼저 내려받으세요.")) return;
    clearInterval(app.timer);
    await api("/api/control/close", {});
    app.report = null;
    goStep(1);
  }));

  // 이미 진행 중인 세션이 있으면(새로고침 등) 이어서 보여 준다.
  try {
    const report = await (await api("/api/status")).json();
    if (report.state && report.state !== "idle") {
      app.mode = report.config.mode;
      app.autoStarted = false;
      app.resultShown = (report.phases_done || []).length > 0;
      $("#runTitle").textContent = MODE_TITLES[app.mode];
      goStep(3);
      startPolling();
      return;
    }
  } catch (err) { showError(err.message); return; }
  goStep(1);
});
