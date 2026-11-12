"""Menu organisation quality on a typical Korean site layout."""
from io import BytesIO

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

from menu_crawler.crawler import EXTRACT, Crawler
from menu_crawler.export import make_workbook
from menu_crawler.models import CrawlConfig

from test_foundation import launch_options

HTML = '''<base href="https://bank.test/">
<div id="skipnav"><a href="#content">본문 바로가기</a><a href="#gnb">주메뉴 바로가기</a></div>
<header><ul class="util"><li><a href="/login">로그인</a></li><li><a href="/join">회원가입</a></li><li><a href="https://en.bank.test/">ENGLISH</a></li></ul>
<nav id="gnb"><ul>
 <li><a href="#">조회</a><ul>
   <li><a href="javascript:void(0)">계좌조회</a><ul><li><a href="/inq/balance">잔액조회</a></li><li><a href="/inq/history">거래내역조회</a></li></ul></li>
   <li><a href="/inq/card">카드조회</a></li></ul></li>
 <li><a href="/transfer">이체</a><ul><li><a href="/transfer/now">즉시이체</a></li><li><a href="/transfer/rsv">예약이체</a></li></ul></li>
 <li><a href="/notice">공지사항</a></li>
</ul></nav>
<div class="m-menu" hidden><nav><ul><li><a href="#">조회</a><ul><li><a href="/inq/card">카드조회</a></li></ul></li></ul></nav></div>
</header>
<div class="lnb"><h2>이체</h2><ul><li><a href="/transfer/now">즉시이체</a></li><li><a href="/transfer/rsv">예약이체</a></li></ul></div>
<main id="content"><a href="/event/1">이벤트 보러가기</a><a href="#top">맨 위로</a></main>
<footer><a href="/terms">이용약관</a><a href="/privacy">개인정보처리방침</a></footer>'''


def collect(html):
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True, **launch_options())
        try:
            page = browser.new_page()
            page.set_content(html)
            data = page.evaluate(EXTRACT, {"customSelector": ""})
        finally:
            browser.close()
    crawler = Crawler(CrawlConfig(url="https://bank.test"))
    for record in data["records"]:
        crawler.add(record, "https://bank.test/", 0)
    return crawler


def test_typical_layout_is_organised():
    crawler = collect(HTML)
    menus = {" > ".join(m.path): m for m in crawler.items.values()}
    # Skip links and in-page anchors are not menus.
    assert not any("바로가기" in k or k == "맨 위로" for k in menus)
    # The LNB copy joins the GNB path instead of becoming a second 1 Depth menu.
    assert "즉시이체" not in menus and "이체 > 즉시이체" in menus
    # Hidden mobile duplicate collapses into the same entry.
    assert sum(1 for k in menus if k.endswith("카드조회")) == 1
    # Utility and footer links are kept but marked as common links.
    assert menus["이용약관"].evidence == "공통 링크" and menus["회원가입"].evidence == "공통 링크"
    assert menus["조회 > 계좌조회 > 거래내역조회"].evidence == "메뉴 DOM"


def test_common_links_are_not_review_work():
    crawler = collect(HTML)
    for m in crawler.items.values():
        crawler.apply_result(m, {"status": "수집됨", "reason": ""})
    wb = load_workbook(BytesIO(make_workbook({**crawler.snapshot(), "state": "complete"})))
    review = [row[1].value for row in wb["미수집_확인필요"].iter_rows(min_row=2)]
    assert "이용약관" not in review and "이벤트 보러가기" in review


def test_lnb_heading_ignored_when_hidden_label():
    html = '''<base href="https://bank.test/"><div class="lnb"><h2 class="blind">주메뉴</h2>
      <ul><li><a href="/a">메뉴A</a></li></ul></div>'''
    assert [m.path for m in collect(html).items.values()] == [["메뉴A"]]


def test_tree_order_groups_late_discoveries():
    crawler = Crawler(CrawlConfig(url="https://bank.test"))
    for path, url in [(["조회", "계좌"], "/a"), (["이체"], "/b"), (["조회", "카드"], "/c"), (["이체", "예약"], "/d")]:
        crawler.add({"path": path, "url": "https://bank.test" + url, "evidence": "메뉴 DOM"}, "https://bank.test/", 0)
    assert [m["path"] for m in crawler.snapshot()["menus"]] == [["조회", "계좌"], ["조회", "카드"], ["이체"], ["이체", "예약"]]
