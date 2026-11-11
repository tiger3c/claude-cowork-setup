from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

from fastapi.testclient import TestClient
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

from menu_crawler.crawler import Crawler, EXTRACT
from menu_crawler.export import make_workbook
from menu_crawler.models import CrawlConfig, Menu, Report
from menu_crawler.urls import canonical_url, public_url, visit_policy
from menu_crawler.web import create_app


def make_report(menus):
    return Report(config=CrawlConfig(url="https://example.test").model_dump(),
                  menus=menus, state="complete").to_dict()


def test_url_keeps_menu_query_and_spa_route():
    assert canonical_url("/view?menuId=10&utm_source=test#/tab", "https://EXAMPLE.test:443") == "https://example.test/view?menuId=10#/tab"
    assert canonical_url("/view#footer", "https://example.test") == "https://example.test/view"
    assert canonical_url("javascript:alert(1)") == ""
    assert canonical_url("https://user:password@example.test") == ""


def test_common_credentials_redacted_from_export_url():
    result = public_url("https://example.test/view?menuId=3&access_token=not-a-real-token&code=example")
    values = parse_qs(urlsplit(result).query)
    assert values == {"menuId": ["3"], "access_token": ["[가림]"], "code": ["[가림]"]}


def test_scope_and_non_navigation_links():
    hosts = {"example.test"}
    assert visit_policy("https://other.test/page", hosts)[0] == "범위 밖"
    assert visit_policy("https://example.test/logout", hosts)[0] == "동작 확인 필요"
    assert visit_policy("https://example.test/file.pdf", hosts)[0] == "파일 링크"
    assert visit_policy("https://example.test/login", hosts)[0] == "로그인 필요"
    assert visit_policy("https://example.test/view?menu=transfer", hosts)[0] == "대기"


def test_fourth_level_is_not_exported_or_visited():
    crawler = Crawler(CrawlConfig(url="https://example.test"))
    for leaf in ("상세 A", "상세 B"):
        crawler.add({"path": ["조회", "계좌", "거래내역", leaf], "url": "https://example.test/" + leaf,
                     "evidence": "메뉴 DOM"}, "https://example.test/", 0)
    menus = crawler.snapshot()["menus"]
    assert len(menus) == 1
    assert menus[0]["path"] == ["조회", "계좌", "거래내역"]
    assert menus[0]["url"] == ""
    assert menus[0]["ready"]
    assert list(crawler.pending) == []
    crawler.add({"path": ["조회", "계좌", "거래내역"], "url": "https://example.test/history",
                 "evidence": "메뉴 DOM"}, "https://example.test/", 0)
    assert len(crawler.items) == 1
    assert next(iter(crawler.items.values())).url.endswith("/history")


def test_one_level_is_valid_and_url_aliases_keep_distinct_paths():
    crawler = Crawler(CrawlConfig(url="https://example.test"))
    for path in (["공지사항"], ["고객센터", "공지사항"]):
        crawler.add({"path": path, "url": "https://example.test/notices", "evidence": "메뉴 DOM"}, "https://example.test/", 0)
    assert len(crawler.items) == 2
    assert len(crawler.pending) == 1
    for menu in crawler.items.values():
        crawler.apply_result(menu, {"status": "수집됨", "reason": ""})
        assert menu.ready


def test_export_preserves_literal_text_and_splits_250_menus():
    menus = [Menu(str(i), [f"메뉴 {i}"], "https://example.test/" + str(i),
                  "https://example.test", "메뉴 DOM", status="수집됨") for i in range(251)]
    menus[0].path = ['=HYPERLINK("https://example.test","테스트")']
    failed = Menu("failed", ["비공개 메뉴"], "https://example.test/private", "https://example.test", "메뉴 DOM", status="로그인 필요", reason="인증 화면 감지")
    menus.append(failed)
    wb = load_workbook(BytesIO(make_workbook(make_report(menus))))
    assert {"수집요약", "전체메뉴", "미수집_확인필요", "방문기록", "진단메뉴_01", "진단메뉴_02"} <= set(wb.sheetnames)
    assert wb["전체메뉴"]["B2"].data_type == "s"
    assert wb["진단메뉴_01"]["H13"].data_type == "s"
    assert wb["진단메뉴_01"].max_column == 257  # H:IW = 250 menu columns.
    assert wb["진단메뉴_02"]["H13"].value == "메뉴 250"
    assert wb["진단메뉴_01"]["H14"].value is None
    assert wb["미수집_확인필요"]["F2"].value == "로그인 필요"
    assert wb["전체메뉴"].freeze_panes == "D2"


def test_explicit_selection_and_three_depth_cells():
    report = make_report([Menu("one", ["원본"], "", "", "일반 링크", hierarchy_review=True)])
    wb = load_workbook(BytesIO(make_workbook(report, [{"id": "one", "depths": ["수정 메뉴", "", ""]}])))
    assert wb["진단메뉴_01"]["H13"].value == "수정 메뉴"
    assert wb["전체메뉴"]["B2"].value == "원본"
    assert "원본" in str(wb["전체메뉴"]["E2"].value)


def test_local_api_session_and_export_route():
    class FakeJob:
        def snapshot(self):
            return make_report([Menu("one", ["공지사항"], "https://example.test/notice", "", "메뉴 DOM", status="수집됨")])

        def close(self):
            pass

    app = create_app("test-session")
    with TestClient(app) as client:
        assert client.get("/api/status").status_code == 403
        headers = {"x-session-token": "test-session"}
        app.state.job = FakeJob()
        assert client.get("/api/status", headers=headers).status_code == 200
        assert client.post("/api/export", headers={**headers, "origin": "https://unrelated.test"}, json={}).status_code == 403
        result = client.post("/api/export", headers=headers, json={})
        assert result.status_code == 200
        assert load_workbook(BytesIO(result.content))["진단메뉴_01"]["H13"].value == "공지사항"


def test_real_browser_reads_hidden_nested_menu():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        try:
            page = browser.new_page()
            page.set_content('''<base href="https://example.test/">
              <nav><ul><li><button>조회</button><ul hidden>
              <li><a href="/accounts">계좌</a><ul><li><span>거래내역</span><ul>
              <li><a href="/detail">상세</a></li></ul></li></ul></li></ul></li>
              <li><a href="/notices">공지사항</a></li></ul></nav>''')
            data = page.evaluate(EXTRACT, {"customSelector": ""})
            assert any(m["path"] == ["조회", "계좌", "거래내역", "상세"] for m in data["records"])
            assert any(m["path"] == ["공지사항"] for m in data["records"])
            crawler = Crawler(CrawlConfig(url="https://example.test"))
            for record in data["records"]:
                crawler.add(record, "https://example.test", 0)
            assert all(len(m.path) <= 3 for m in crawler.items.values())
            assert not any(url.endswith("/detail") for url, _ in crawler.pending)
        finally:
            browser.close()
