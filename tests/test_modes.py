"""Collection modes: before login, after login, and the before/after comparison."""
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from menu_crawler.crawler import Crawler
from menu_crawler.export import make_workbook
from menu_crawler.models import CrawlConfig
from menu_crawler.web import create_app

from test_foundation import launch_options


def add(crawler, path, url, phase):
    crawler.phase = phase
    crawler.add({"path": path, "url": "https://example.test" + url, "evidence": "메뉴 DOM"}, "https://example.test/", 0)


def by_path(snapshot):
    return {" > ".join(m["path"]): m for m in snapshot["menus"]}


def test_compare_mode_marks_menus_seen_only_after_login():
    crawler = Crawler(CrawlConfig(url="https://example.test", mode="compare"))
    add(crawler, ["고객센터", "공지사항"], "/notice", "before")
    add(crawler, ["고객센터", "공지사항"], "/notice", "after")
    add(crawler, ["계좌", "조회", "거래내역"], "/history", "after")
    menus = by_path(crawler.snapshot())
    assert menus["고객센터 > 공지사항"]["visibility"] == "로그인 전 공개"
    assert menus["계좌 > 조회 > 거래내역"]["visibility"] == "로그인 후 전용"
    for m in crawler.items.values():
        crawler.apply_result(m, {"status": "수집됨", "reason": ""})
    report = {**crawler.snapshot(), "state": "complete"}
    wb = load_workbook(BytesIO(make_workbook(report)))
    only = wb["로그인후_전용메뉴"]
    assert only.max_row == 2 and only["D2"].value == "거래내역"
    assert wb["전체메뉴"]["O1"].value == "노출 구분"
    assert any(row[0].value == "수집 방식" and row[1].value == "로그인 전·후 비교 수집" for row in wb["수집요약"].iter_rows())


def test_other_modes_do_not_claim_visibility():
    for mode in ("public", "login"):
        crawler = Crawler(CrawlConfig(url="https://example.test", mode=mode))
        add(crawler, ["공지사항"], "/notice", crawler.phase)
        assert crawler.snapshot()["menus"][0]["visibility"] == ""
        wb = load_workbook(BytesIO(make_workbook({**crawler.snapshot(), "state": "complete"})))
        assert "로그인후_전용메뉴" not in wb.sheetnames
    assert Crawler(CrawlConfig(url="https://example.test", mode="login")).phase == "after"


def test_login_url_is_validated():
    with pytest.raises(ValueError):
        CrawlConfig(url="https://example.test", mode="login", login_url="javascript:alert(1)")
    assert CrawlConfig(url="https://example.test", login_url=" ").login_url == ""


class FakeJob:
    def __init__(self, mode):
        self.config = CrawlConfig(url="https://example.test", mode=mode)
        self.calls = []

    def command(self, name, **args):
        self.calls.append((name, args))

    def snapshot(self):
        return {"state": "ready"}

    def close(self):
        pass


def test_api_routes_phases_by_mode_and_never_echoes_password():
    app = create_app("t")
    headers = {"x-session-token": "t"}
    with TestClient(app) as client:
        assert client.get("/").status_code == 200
        assert client.get("/static/app.js").status_code == 200
        app.state.job = FakeJob("public")
        assert client.post("/api/control/crawl", headers=headers, json={"phase": "after"}).status_code == 409
        assert client.post("/api/control/login", headers=headers, json={}).status_code == 409
        assert client.post("/api/login/fill", headers=headers, json={"password": "x"}).status_code == 409
        assert client.post("/api/control/crawl", headers=headers, json={}).status_code == 200
        assert app.state.job.calls[-1] == ("crawl", {"phase": "before"})

        app.state.job = FakeJob("compare")
        assert client.post("/api/control/crawl", headers=headers, json={"phase": "after"}).status_code == 200
        assert client.post("/api/control/login", headers=headers, json={}).status_code == 200
        assert app.state.job.calls[-1] == ("goto", {"url": "https://example.test/"})
        secret = "not-a-real-password-" + "x" * 600
        result = client.post("/api/login/fill", headers=headers, json={"username": "u", "password": secret})
        assert result.status_code == 422 and "not-a-real-password" not in result.text
        assert client.post("/api/login/fill", headers=headers, json={"username": "u", "password": "pw"}).status_code == 200
        assert app.state.job.calls[-1] == ("fill", {"username": "u", "password": "pw"})


# ---------- end-to-end against a local test site ----------

PAGE = """<!doctype html><html><head><title>{title}</title></head><body>
<nav><ul>
  <li><span>고객센터</span><ul><li><a href="/notice">공지사항</a></li></ul></li>
  <li><a href="/mypage">마이페이지</a></li>
  {member}
</ul></nav><main>{body}</main></body></html>"""
MEMBER = """<li><span>계좌</span><ul><li><span>조회</span><ul><li><a href="/history">거래내역</a></li></ul></li></ul></li>
  <li><a href="/logout">로그아웃</a></li>"""
GUEST = '<li><a href="/login">로그인</a></li>'
LOGIN_FORM = """<form method="post" action="/login"><input name="userid" type="text">
<input name="pw" type="password"><button type="submit">로그인</button></form>"""


class Site(BaseHTTPRequestHandler):
    logins = []

    def log_message(self, *args):
        pass

    def authed(self):
        return "sid=ok" in (self.headers.get("Cookie") or "")

    def send_page(self, title, body, status=200):
        html = PAGE.format(title=title, body=body, member=MEMBER if self.authed() else GUEST).encode()
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(html)

    def do_GET(self):
        if self.path == "/logout":
            raise AssertionError("logout must never be visited automatically")
        if self.path == "/login":
            return self.send_page("로그인", LOGIN_FORM)
        if self.path in {"/mypage", "/history"} and not self.authed():
            return self.send_page("로그인", "로그인이 필요합니다." + LOGIN_FORM)
        self.send_page(self.path, "본문")

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        Site.logins.append(self.rfile.read(length).decode())
        self.send_response(302)
        self.send_header("Set-Cookie", "sid=ok; Path=/")
        self.send_header("Location", "/")
        self.end_headers()


@pytest.fixture()
def site():
    server = ThreadingHTTPServer(("127.0.0.1", 0), Site)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{server.server_address[1]}"
    server.shutdown()


def test_compare_flow_against_local_login_site(site, monkeypatch):
    if launch_options():
        monkeypatch.setenv("MENU_CRAWLER_BROWSER_PATH", launch_options()["executable_path"])
    config = CrawlConfig(url=site, mode="compare", login_url=site + "/login", browser="chromium",
                         headless=True, wait_ms=0, delay_ms=0)
    engine = Crawler(config)
    try:
        engine.open()
        engine.crawl("before")
        before = by_path(engine.snapshot())
        assert "계좌 > 조회 > 거래내역" not in before
        assert before["마이페이지"]["status"] == "로그인 필요"

        engine.goto(config.login_url)
        engine.fill_login("tester", "not-a-real-password")
        assert "not-a-real-password" not in str(engine.snapshot())
        engine.page.click("button[type=submit]")  # the user presses the login button
        engine.page.wait_for_url(site + "/")
        assert Site.logins[-1] == "userid=tester&pw=not-a-real-password"

        engine.crawl("after")
        report = engine.snapshot()
        menus = by_path(report)
        assert menus["계좌 > 조회 > 거래내역"]["visibility"] == "로그인 후 전용"
        assert menus["계좌 > 조회 > 거래내역"]["status"] == "수집됨"
        assert menus["고객센터 > 공지사항"]["visibility"] == "로그인 전 공개"
        assert menus["마이페이지"]["status"] == "수집됨"
        assert menus["마이페이지"]["status_before"] == "로그인 필요"
        assert menus["로그아웃"]["status"] == "동작 확인 필요"
        mypage = [x for x in report["issues"] if x["url"].endswith("/mypage")]
        assert mypage and all(x["phase"] == "로그인 전" and x["resolved"] == "로그인 후 수집됨" for x in mypage)
        assert report["phases_done"] == ["before", "after"]
        assert not report["login_check"]
        assert "not-a-real-password" not in str(report)
        wb = load_workbook(BytesIO(make_workbook(report)))
        assert wb["로그인후_전용메뉴"].max_row >= 2
    finally:
        engine.close()
