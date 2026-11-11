import copy
import hashlib
import json
import queue
import threading
import time
from collections import deque
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlsplit

from playwright.sync_api import Error as BrowserError
from playwright.sync_api import sync_playwright

from .models import CrawlConfig, Menu, Report
from .urls import AUTH, canonical_url, public_url, visit_policy

EXTRACT = Path(__file__).with_name("extract.js").read_text(encoding="utf-8")


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
        self.issue_keys: set[tuple] = set()
        self.pw = self.browser = self.context = self.page = None

    def snapshot(self):
        self.report.menus = list(self.items.values())
        result = self.report.to_dict()
        result["config"]["url"] = public_url(result["config"]["url"])
        result["current_url"] = public_url(result["current_url"])
        for item in result["menus"]:
            for key in ["url", "source_url", "final_url"]:
                item[key] = public_url(item[key])
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
        key = kind, url, reason
        if key not in self.issue_keys:
            self.issue_keys.add(key)
            self.report.issues.append({"kind": kind, "url": url, "reason": reason})

    def open(self):
        self.pw = sync_playwright().start()
        choices = ["msedge", "chrome", "chromium"] if self.config.browser == "auto" else [self.config.browser]
        for channel in choices:
            try:
                self.browser = self.pw.chromium.launch(headless=self.config.headless,
                    **({"channel": channel} if channel != "chromium" else {}))
                break
            except BrowserError:
                continue
        if not self.browser:
            raise RuntimeError("브라우저를 열지 못했습니다. Edge/Chrome을 설치하거나 브라우저설치.bat를 실행하세요. 회사 브라우저 정책도 확인하세요.")
        self.context = self.browser.new_context(accept_downloads=False, viewport={"width": 1440, "height": 1000})
        self.context.on("page", lambda p: p.on("dialog", lambda dialog: dialog.dismiss()))
        self.page = self.context.new_page()
        self.page.set_default_timeout(self.config.timeout_ms)
        try:
            self.page.goto(self.config.url, wait_until="domcontentloaded", timeout=self.config.timeout_ms)
        except BrowserError as exc:
            self.issue("시작 페이지 접속 실패", self.config.url, self.error_text(exc))
        self.emit("ready", "브라우저에서 필요한 로그인을 마친 후 수집 시작을 누르세요. 공개 사이트는 바로 시작해도 됩니다.")

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
        if kind == "group" and any(m.path == path and m.kind == "link" for m in self.items.values()):
            return
        if kind == "link":
            for key, m in list(self.items.items()):
                if m.path == path and m.kind == "group":
                    del self.items[key]
        # Prefer an actual navigation hierarchy over a duplicate generic page link.
        if record.get("evidence") == "일반 링크" and any(m.url == url and m.evidence == "메뉴 DOM" for m in self.items.values()):
            return
        if record.get("evidence") == "메뉴 DOM" and url:
            for key, m in list(self.items.items()):
                if m.url == url and m.evidence == "일반 링크":
                    del self.items[key]
        key = tuple(path), url, source if kind == "dynamic" or review else ""
        if key in self.items:
            return
        status, reason = visit_policy(url, self.hosts)
        if kind == "group":
            status, reason = "분류 메뉴", "하위 메뉴를 묶는 제목이며 독립적인 페이지 주소는 없습니다."
        ident = hashlib.sha256(json.dumps(key, ensure_ascii=False).encode()).hexdigest()[:16]
        menu = Menu(ident, path, url, source, record.get("evidence", "확인 필요"), kind, status, reason, hierarchy_review=review)
        if url in self.results and status == "대기":
            self.apply_result(menu, self.results[url])
        self.items[key] = menu
        if menu.status == "대기":
            self.enqueue(url, hops + 1)

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

    def capture(self):
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
        self.emit("ready", f"열린 화면 {count}개에서 메뉴를 추가했습니다. 수집 시작을 누르면 접속 상태도 확인합니다.")

    def crawl(self):
        self.stop.clear()
        self.report.pages = []
        self.report.visited_count = 0
        self.report.limit_reached = False
        self.report.finished_at = ""
        self.results.clear()
        self.pending.clear()
        self.scheduled.clear()
        self.emit("crawling", "메뉴와 페이지를 확인하고 있습니다.")
        for item in self.items.values():
            if item.kind == "link" and len(item.path) <= 3:
                item.status, item.reason = visit_policy(item.url, self.hosts)
                self.enqueue(item.url, 0)
        self.enqueue(self.config.url, 0)
        for page in list(self.context.pages):
            if not page.is_closed():
                self.harvest(page, 0, manual=True)
                self.enqueue(canonical_url(page.url), 0)
        deadline = time.monotonic() + self.config.max_seconds
        end_reason = ""
        while self.pending:
            if self.stop.is_set():
                end_reason = "사용자가 수집을 중지했습니다."
                break
            if self.report.visited_count >= self.config.max_pages:
                end_reason = "최대 방문 페이지 수에 도달했습니다."
                break
            if time.monotonic() >= deadline:
                end_reason = "최대 수집 시간에 도달했습니다."
                break
            url, hops = self.pending.popleft()
            self.report.current_url = url
            self.emit()
            result = {"url": url, "final_url": "", "title": "", "http_status": None, "status": "수집됨", "reason": ""}
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
        self.report.finished_at = datetime.now(timezone.utc).isoformat()
        self.report.current_url = ""
        self.emit("stopped" if self.stop.is_set() else "complete",
                  "수집을 종료했습니다. 미수집·확인 필요 항목을 검토한 뒤 엑셀을 내려받으세요.")


class Job:
    def __init__(self, config):
        self.stop = threading.Event()
        self.commands = queue.Queue()
        self.lock = threading.Lock()
        self.data = Report(config=config.model_dump()).to_dict()
        self.thread = threading.Thread(target=self.run, args=(config,), daemon=True)
        self.thread.start()

    def publish(self, data):
        with self.lock:
            self.data = copy.deepcopy(data)

    def snapshot(self):
        with self.lock:
            return copy.deepcopy(self.data)

    def command(self, name):
        with self.lock:
            if self.data["state"] not in {"ready", "complete", "stopped"}:
                raise ValueError("현재 작업을 마치거나 중지한 후 실행하세요.")
            self.data["state"] = "crawling" if name == "crawl" else "capturing"
        self.commands.put(name)

    def close(self):
        self.stop.set()
        self.commands.put("close")

    def run(self, config):
        engine = Crawler(config, self.publish, self.stop)
        try:
            engine.open()
            while True:
                command = self.commands.get()
                if command == "close":
                    break
                if command == "crawl":
                    engine.crawl()
                elif command == "capture":
                    engine.capture()
        except Exception as exc:
            message = str(exc) if isinstance(exc, RuntimeError) else Crawler.error_text(exc)
            engine.issue("실행 오류", config.url, message)
            engine.emit("error", message)
        finally:
            engine.close()
