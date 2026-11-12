import copy
import hashlib
import json
import os
import queue
import threading
import time
from collections import deque
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlsplit

from playwright.sync_api import Error as BrowserError
from playwright.sync_api import sync_playwright

from .models import PHASES, CrawlConfig, Menu, Report
from .urls import AUTH, canonical_url, page_shape, public_url, visit_policy

EXTRACT = Path(__file__).with_name("extract.js").read_text(encoding="utf-8")


SAME_PAGE_VISITS = 3     # visits of one page whose links differ only in a token / timestamp

class Crawler:
    """Owned by one thread, including every Playwright call."""

    def __init__(self, config: CrawlConfig, publish=lambda report: None, stop=None):
        self.config = config
        self.report = Report(config=config.model_dump())
        self.publish = publish
        self.stop = stop or threading.Event()
        self.hosts = {urlsplit(config.url).netloc.lower(), *config.extra_hosts}
        self.items: dict[tuple, Menu] = {}
        self.results: dict[str, dict] = {}
        self.pending: deque[tuple[str, int]] = deque()
        self.scheduled: set[str] = set()
        self.shapes: dict[str, int] = {}
        self.issue_keys: set[tuple] = set()
        self.phase = "after" if config.mode == "login" else "before"
        self.pw = self.browser = self.context = self.page = None

    def ordered_menus(self):
        """Tree order: menus stay grouped under their 1 Depth, then 2 Depth, in first-seen order."""
        rank: dict[tuple, int] = {}
        for menu in self.items.values():
            for depth in range(1, len(menu.path) + 1):
                rank.setdefault(tuple(menu.path[:depth]), len(rank))
        return sorted(self.items.values(), key=lambda m: [rank[tuple(m.path[:d])] for d in range(1, len(m.path) + 1)])

    def snapshot(self):
        self.report.menus = self.ordered_menus()
        result = self.report.to_dict()
        result["config"]["url"] = public_url(result["config"]["url"])
        result["current_url"] = public_url(result["current_url"])
        for item in result["menus"]:
            for key in ["url", "source_url", "final_url"]:
                item[key] = public_url(item[key])
            if self.config.mode != "compare":
                item["visibility"] = ""
        for item in result["pages"]:
            for key in ["url", "final_url"]:
                item[key] = public_url(item.get(key, ""))
        for item in result["issues"]:
            item["url"] = public_url(item.get("url", ""))
        return result

    def emit(self, state=None, message=None):
        if state:
            self.report.state = state
        if message:
            self.report.message = message
        self.publish(self.snapshot())

    def issue(self, kind, url, reason):
        key = kind, url, reason, self.phase
        if key not in self.issue_keys:
            self.issue_keys.add(key)
            self.report.issues.append({"kind": kind, "url": url, "reason": reason,
                                       "phase": PHASES[self.phase], "resolved": ""})

    def open(self):
        self.pw = sync_playwright().start()
        choices = ["msedge", "chrome", "chromium"] if self.config.browser == "auto" else [self.config.browser]
        # Optional: a specific Chromium binary (offline PCs, test machines).
        custom = os.environ.get("MENU_CRAWLER_BROWSER_PATH", "")
        for channel in choices:
            try:
                extra = {"channel": channel} if channel != "chromium" else {"executable_path": custom} if custom else {}
                self.browser = self.pw.chromium.launch(headless=self.config.headless, **extra)
                break
            except BrowserError:
                continue
        if not self.browser:
            raise RuntimeError("브라우저를 열지 못했습니다. Edge/Chrome을 설치하거나 install_browser.bat를 실행하세요. 회사 브라우저 정책도 확인하세요.")
        # Test/dev servers of a diagnosis target often use self-signed certificates.
        self.context = self.browser.new_context(accept_downloads=False, ignore_https_errors=True, viewport={"width": 1440, "height": 1000})
        self.context.on("page", lambda p: p.on("dialog", lambda dialog: dialog.dismiss()))
        self.page = self.context.new_page()
        self.page.set_default_timeout(self.config.timeout_ms)
        first = self.config.login_url if self.config.mode == "login" and self.config.login_url else self.config.url
        try:
            self.page.goto(first, wait_until="domcontentloaded", timeout=self.config.timeout_ms)
        except BrowserError as exc:
            self.issue("시작 페이지 접속 실패", first, self.error_text(exc))
        self.emit("ready", {
            "public": "브라우저가 열렸습니다. 로그인 없이 공개 메뉴를 수집합니다.",
            "login": "열린 브라우저 창에서 로그인을 마친 뒤 로그인 후 수집을 시작하세요.",
            "compare": "1단계: 로그인하지 않은 상태로 공개 메뉴를 먼저 수집합니다.",
        }[self.config.mode])

    @staticmethod
    def error_text(exc):
        text = str(exc)
        if "ERR_CERT" in text or "SSL" in text:
            return "인증서·보안 연결 오류입니다. 브라우저에서 접속 상태를 확인하세요."
        if "Timeout" in text:
            return "페이지 응답 대기 시간을 초과했습니다."
        if "closed" in text.lower():
            return "수집 브라우저가 닫혔습니다. 브라우저 열기부터 다시 시작하세요."
        return "페이지 접속 또는 화면 읽기에 실패했습니다. 브라우저에서 해당 주소를 확인하세요."

    def close(self):
        try:
            if self.browser:
                self.browser.close()
        finally:
            if self.pw:
                self.pw.stop()

    def enqueue(self, url, hops):
        if not url or url in self.scheduled or visit_policy(url, self.hosts)[0] != "대기":
            return
        if hops > self.config.max_hops:
            self.issue("탐색 이동 제한", url, "최대 이동 횟수 때문에 이 주소는 방문하지 못했습니다.")
            return
        shape = page_shape(url)
        seen = self.shapes.get(shape, 0)
        self.shapes[shape] = seen + 1
        if seen >= SAME_PAGE_VISITS:     # token / timestamp in the link: the same page again and again
            if seen == SAME_PAGE_VISITS:
                self.issue("같은 페이지 반복", url, f"주소의 토큰·시각 값만 바뀐 같은 페이지라 {SAME_PAGE_VISITS}번까지만 방문했습니다.")
            return
        self.scheduled.add(url)
        self.pending.append((url, hops))

    def add(self, record, source, hops):
        path = [" ".join(str(x).split())[:240] for x in record.get("path", []) if str(x).strip()]
        if not path:
            return
        url = canonical_url(record.get("url", ""), source) if record.get("url") else ""
        kind = record.get("kind", "link")
        review = bool(record.get("hierarchy_review"))
        if len(path) > 3:
            # The requested inventory ends at level three. Do not visit or export
            # deeper destinations, or mislabel a level-four URL as level three.
            path, url, kind = path[:3], "", "group"
        inherited: set[str] = set()
        if kind == "group":
            links = [m for m in self.items.values() if m.path == path and m.kind == "link"]
            if links:
                self.mark_seen(links)
                return
        if kind == "link":
            for key, m in list(self.items.items()):
                if m.path == path and m.kind == "group":
                    inherited.update(m.seen)
                    del self.items[key]
        # Prefer an actual navigation hierarchy over a duplicate generic page link.
        if record.get("evidence") in {"일반 링크", "공통 링크"}:
            menus = [m for m in self.items.values() if m.url == url and m.evidence == "메뉴 DOM"]
            if menus:
                self.mark_seen(menus)
                return
        if record.get("evidence") == "메뉴 DOM" and url:
            for key, m in list(self.items.items()):
                if m.url == url and m.evidence in {"일반 링크", "공통 링크"}:
                    del self.items[key]
        key = tuple(path), page_shape(url) if url else "", source if kind == "dynamic" or review else ""
        if key in self.items:
            self.mark_seen([self.items[key]])
            return
        status, reason = visit_policy(url, self.hosts)
        if kind == "group":
            status, reason = "분류 메뉴", "하위 메뉴를 묶는 제목이며 독립적인 페이지 주소는 없습니다."
        ident = hashlib.sha256(json.dumps(key, ensure_ascii=False).encode()).hexdigest()[:16]
        menu = Menu(ident, path, url, source, record.get("evidence", "확인 필요"), kind, status, reason, hierarchy_review=review,
                    seen=sorted(inherited | {self.phase}))
        if url in self.results and status == "대기":
            self.apply_result(menu, self.results[url])
        self.items[key] = menu
        if menu.status == "대기":
            self.enqueue(url, hops + 1)

    def mark_seen(self, menus):
        for menu in menus:
            if self.phase not in menu.seen:
                menu.seen = sorted({*menu.seen, self.phase})

    @staticmethod
    def apply_result(menu, result):
        menu.status = result["status"]
        menu.reason = result["reason"]
        menu.http_status = result.get("http_status")
        menu.title = result.get("title", "")
        menu.final_url = result.get("final_url", "")

    def harvest(self, page, hops, manual=False):
        main_data = None
        for frame in list(page.frames):
            raw = frame.url
            if frame != page.main_frame and not raw.lower().startswith(("http:", "https:")):
                continue  # about:blank, srcdoc, chrome-error: no menu source of its own
            source = canonical_url(raw) or canonical_url(page.url)
            if not source or urlsplit(source).netloc.lower() not in self.hosts:
                self.issue("프레임 범위 밖", raw, "이 프레임의 호스트가 수집 범위에 포함되지 않습니다.")
                continue
            try:
                data = frame.evaluate(EXTRACT, {"customSelector": self.config.menu_selector})
            except BrowserError:
                self.issue("화면 읽기 실패", source, "프레임이 바뀌었거나 메뉴 선택자가 잘못됐습니다. 해당 화면을 확인하세요.")
                continue
            if frame == page.main_frame:
                main_data = data
            for item in data["records"]:
                self.add(item, source, hops)
            if data["unnamed"]:
                self.issue("이름 없는 링크", source, f"메뉴명을 읽을 수 없는 링크 {data['unnamed']}개가 있습니다.")
            if manual:
                crumbs = data["breadcrumbs"]
                if crumbs:
                    self.add({"path": crumbs, "url": source, "evidence": "현재 화면 경로", "hierarchy_review": False}, source, hops)
                elif not any(m.url == source for m in self.items.values()):
                    self.add({"path": [data["title"] or "이름 확인 필요"], "url": source,
                              "evidence": "현재 화면 제목", "hierarchy_review": True}, source, hops)
        return main_data

    def hover_menus(self, page, hops):
        """Hover only; never submit forms or click arbitrary application actions."""
        selectors = self.config.menu_selector or "nav,[role=navigation],.gnb,#gnb,.lnb,#lnb"
        try:
            navs = page.locator(selectors)
            controls = navs.locator("[aria-haspopup],button[aria-expanded]")
            count = controls.count()
            if count > 12:
                self.issue("동적 메뉴 확인", page.url, "자동 펼치기 확인은 첫 12개 메뉴에 적용됐습니다. 나머지는 직접 열어 현재 화면 추가를 사용하세요.")
            for i in range(min(count, 12)):
                if self.stop.is_set():
                    break
                item = controls.nth(i)
                if item.is_visible():
                    item.hover(timeout=1500)
                    page.wait_for_timeout(min(self.config.wait_ms, 300))
                    self.harvest(page, hops)
        except BrowserError:
            self.issue("동적 메뉴 확인", page.url, "일부 메뉴가 마우스 이동만으로 열리지 않았습니다. 직접 열어 현재 화면 추가를 사용하세요.")

    def capture(self, phase=None):
        if phase:
            self.phase = phase
        count = 0
        for page in list(self.context.pages):
            if page.is_closed():
                continue
            url = canonical_url(page.url)
            if not url or urlsplit(url).netloc.lower() not in self.hosts:
                self.issue("현재 화면 범위 밖", page.url, "로그인 후 대상 사이트로 돌아오거나 허용 호스트에 추가하세요.")
                continue
            page.wait_for_timeout(self.config.wait_ms)
            self.harvest(page, 0, manual=True)
            self.hover_menus(page, 0)
            count += 1
        self.emit("ready", f"열린 화면 {count}개에서 메뉴를 추가했습니다({PHASES[self.phase]}). 수집을 시작하면 접속 상태도 확인합니다.")

    def goto(self, url):
        """Bring the login screen to the front of the collection browser."""
        try:
            if self.page.is_closed():
                self.page = self.context.new_page()
            self.page.goto(url, wait_until="domcontentloaded", timeout=self.config.timeout_ms)
            self.page.bring_to_front()
            self.emit(None, "로그인 화면을 열었습니다. 브라우저 창에서 로그인을 마친 뒤 로그인 후 수집을 시작하세요.")
        except BrowserError as exc:
            self.issue("로그인 화면 접속 실패", url, self.error_text(exc))
            self.emit(None, "로그인 화면을 열지 못했습니다. 브라우저 창에서 직접 로그인 화면으로 이동하세요.")

    def fill_login(self, username, password):
        """Type credentials into the visible login form. The user presses the login button.

        Values are used once and never stored in the report, logs or exports.
        """
        for page in reversed([p for p in self.context.pages if not p.is_closed()]):
            for frame in page.frames:
                try:
                    fields = frame.locator('input[type="password"]:visible')
                    if not fields.count():
                        continue
                    pw_field = fields.first
                    message = "아이디와 비밀번호를 입력했습니다."
                    if username:
                        user_field = self.username_field(pw_field)
                        if user_field is None:
                            message = "아이디 입력란을 찾지 못해 비밀번호만 입력했습니다. 아이디는 브라우저 창에서 직접 입력하세요."
                        else:
                            user_field.fill(username)
                    pw_field.fill(password)
                    page.bring_to_front()
                    self.report.login_check = ""
                    self.emit(None, message + " 보안문자·추가 인증이 있으면 완료하고, 로그인 버튼은 직접 누르세요.")
                    return
                except BrowserError:
                    continue
        self.emit(None, "비밀번호 입력란을 찾지 못했습니다. 브라우저 창에서 로그인 화면을 연 뒤 다시 시도하거나 직접 입력하세요.")

    @staticmethod
    def username_field(pw_field):
        """Nearest visible text-like input before the password, inside the same form if there is one."""
        kinds = 'not(@type) or @type="text" or @type="email" or @type="tel" or @type="number"'
        for xpath in (f'xpath=ancestor::form[1]//input[{kinds}]', f'xpath=preceding::input[{kinds}]'):
            candidates = pw_field.locator(xpath)
            pw_box = pw_field.bounding_box() or {"y": 0}
            found = None
            for i in range(min(candidates.count(), 30)):
                item = candidates.nth(i)
                if not item.is_visible():
                    continue
                box = item.bounding_box()
                # Keep the last visible field that sits before (above or left of) the password.
                if box and box["y"] <= pw_box["y"] + 5:
                    found = item
            if found is not None:
                return found
        return None

    def login_screen_visible(self):
        for page in self.context.pages:
            if page.is_closed():
                continue
            try:
                if page.locator('input[type="password"]:visible').count():
                    return True
            except BrowserError:
                continue
        return False

    def crawl(self, phase=None):
        if phase:
            self.phase = phase
        self.stop.clear()
        # A re-crawl replaces the warnings of the same phase; earlier phases stay as history.
        label = PHASES[self.phase]
        self.report.issues = [x for x in self.report.issues if x.get("phase") != label]
        self.issue_keys = {k for k in self.issue_keys if k[3] != self.phase}
        self.report.phase = self.phase
        self.report.login_check = ""
        if self.phase == "after" and self.login_screen_visible():
            self.report.login_check = "열린 화면에 아직 비밀번호 입력란이 보입니다. 로그인이 끝났는지 확인하세요."
            self.issue("로그인 상태 확인", canonical_url(self.page.url), self.report.login_check)
        self.report.pages = [x for x in self.report.pages if x.get("phase") != label]
        self.report.visited_count = 0
        self.report.limit_reached = False
        self.report.finished_at = ""
        self.results.clear()
        self.pending.clear()
        self.scheduled.clear()
        self.shapes.clear()
        if self.phase == "after" and "before" in self.report.phases_done:
            for item in self.items.values():
                if not item.status_before and "before" in item.seen:
                    item.status_before = item.status
        self.emit("crawling", f"{label} 상태로 메뉴와 페이지를 확인하고 있습니다.")
        for item in self.items.values():
            if item.kind == "link" and len(item.path) <= 3:
                item.status, item.reason = visit_policy(item.url, self.hosts)
                self.enqueue(item.url, 0)
        self.enqueue(self.config.url, 0)
        for page in list(self.context.pages):
            if not page.is_closed():
                self.harvest(page, 0, manual=True)
                self.enqueue(canonical_url(page.url), 0)
        deadline = time.monotonic() + self.config.max_seconds if self.config.max_seconds else None
        end_reason = ""
        while self.pending:
            if self.stop.is_set():
                end_reason = "사용자가 수집을 중지했습니다."
                break
            if self.config.max_pages and self.report.visited_count >= self.config.max_pages:
                end_reason = "최대 방문 페이지 수에 도달했습니다."
                break
            if deadline is not None and time.monotonic() >= deadline:
                end_reason = "최대 수집 시간에 도달했습니다."
                break
            url, hops = self.pending.popleft()
            self.report.current_url = url
            self.emit()
            result = {"url": url, "final_url": "", "title": "", "http_status": None, "status": "수집됨", "reason": "", "phase": label}
            try:
                if self.page.is_closed():
                    self.page = self.context.new_page()
                response = self.page.goto(url, wait_until="domcontentloaded", timeout=self.config.timeout_ms)
                self.page.wait_for_timeout(self.config.wait_ms)
                result.update(final_url=canonical_url(self.page.url), http_status=response.status if response else None)
                final_host = urlsplit(result["final_url"]).netloc.lower()
                if final_host not in self.hosts:
                    result.update(status="범위 밖 이동", reason="리다이렉트된 호스트가 수집 범위 밖입니다. 로그인 또는 호스트 설정을 확인하세요.")
                else:
                    data = self.harvest(self.page, hops)
                    result["title"] = data["title"] if data else ""
                    code = result["http_status"]
                    if code in {401, 403}:
                        result.update(status="로그인·권한 확인", reason=f"HTTP {code}: 로그인 또는 접근 권한 확인이 필요합니다.")
                    elif code and code >= 400:
                        result.update(status="접속 실패", reason=f"HTTP {code}: 페이지를 정상적으로 읽지 못했습니다.")
                    elif AUTH.search(result["final_url"]) or (data and (data["password"] or data["loginText"])):
                        result.update(status="로그인 필요", reason="로그인·인증 화면이 감지됐습니다. 인증 후 재수집하세요.")
                    elif data is None:
                        result.update(status="화면 읽기 실패", reason="메뉴 구조를 읽지 못했습니다.")
                    if hops == 0 and result["status"] == "수집됨":
                        self.hover_menus(self.page, hops)
            except BrowserError as exc:
                result.update(status="접속 실패", reason=self.error_text(exc))
            self.results[url] = result
            for menu in self.items.values():
                if menu.url == url and menu.status != "분류 메뉴":
                    self.apply_result(menu, result)
            self.report.pages.append(result)
            self.report.visited_count += 1
            if result["status"] != "수집됨":
                self.issue(result["status"], url, result["reason"])
            self.emit()
            if self.config.delay_ms:
                self.stop.wait(self.config.delay_ms / 1000)
        if end_reason:
            self.report.limit_reached = True
            self.issue("수집 미완료", self.config.url, end_reason)
        for menu in self.items.values():
            if menu.status == "대기":
                menu.status = "미방문"
                menu.reason = end_reason or "탐색 이동 제한 등으로 방문하지 못했습니다."
        if self.phase == "after":
            visited = [p for p in self.report.pages if p.get("phase") == label]
            login_pages = sum(1 for p in visited if p["status"] in {"로그인 필요", "로그인·권한 확인"})
            if visited and login_pages / len(visited) >= 0.3:
                self.issue("로그인 세션 확인", self.config.url,
                           f"로그인 후 수집인데 방문 페이지 {len(visited)}개 중 {login_pages}개가 로그인·권한 화면입니다. "
                           "로그인이 풀렸거나 세션이 만료됐을 수 있습니다. 다시 로그인한 뒤 재수집하세요.")
            for item in self.report.issues:
                if item.get("phase") == PHASES["before"] and self.results.get(item.get("url"), {}).get("status") == "수집됨":
                    item["resolved"] = "로그인 후 수집됨"
        if self.phase not in self.report.phases_done:
            self.report.phases_done.append(self.phase)
        self.report.finished_at = datetime.now(timezone.utc).isoformat()
        self.report.current_url = ""
        if self.config.mode == "compare" and self.phase == "before" and not self.stop.is_set():
            message = "1단계(로그인 전) 수집을 마쳤습니다. 다음은 로그인 단계입니다."
        else:
            message = "수집을 종료했습니다. 미수집·확인 필요 항목은 결과 엑셀에서 확인하세요."
        self.emit("stopped" if self.stop.is_set() else "complete", message)


class Job:
    def __init__(self, config):
        self.config = config
        self.stop = threading.Event()
        self.commands = queue.Queue()
        self.lock = threading.Lock()
        self.issued = 0  # commands queued so far; command() returns this as a ticket
        self.completed = 0  # commands finished so far
        self.data = Report(config=config.model_dump()).to_dict()
        self.thread = threading.Thread(target=self.run, args=(config,), daemon=True)
        self.thread.start()

    def publish(self, data):
        with self.lock:
            self.data = copy.deepcopy(data)

    def snapshot(self):
        with self.lock:
            return copy.deepcopy(self.data)

    def command(self, name, **args):
        with self.lock:
            if self.data["state"] not in {"ready", "complete", "stopped"}:
                raise ValueError("현재 작업을 마치거나 중지한 후 실행하세요.")
            if name in {"crawl", "capture"}:
                self.data["state"] = "crawling" if name == "crawl" else "capturing"
            self.issued += 1
            ticket = self.issued
            self.commands.put((name, args))
        return ticket

    def wait_done(self, ticket, timeout=30):
        end = time.monotonic() + timeout
        while self.completed < ticket and self.thread.is_alive() and time.monotonic() < end:
            time.sleep(0.1)
        return self.snapshot()

    def close(self):
        self.stop.set()
        self.commands.put(("close", {}))

    def run(self, config):
        engine = Crawler(config, self.publish, self.stop)
        try:
            engine.open()
            while True:
                command, args = self.commands.get()
                if command == "close":
                    break
                if command == "crawl":
                    engine.crawl(args.get("phase"))
                elif command == "capture":
                    engine.capture(args.get("phase"))
                elif command == "goto":
                    engine.goto(args["url"])
                elif command == "fill":
                    engine.fill_login(args.get("username", ""), args.get("password", ""))
                args.clear()
                self.completed += 1
        except Exception as exc:
            message = str(exc) if isinstance(exc, RuntimeError) else Crawler.error_text(exc)
            engine.issue("실행 오류", config.url, message)
            engine.emit("error", message)
        finally:
            engine.close()
