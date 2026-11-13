"""The whole interactive flow: work Excel -> targets -> crawl -> fill, driven by scripted answers."""
import json
import os
import shutil
import subprocess
import urllib.request
from pathlib import Path

import pytest
from openpyxl import load_workbook

from menu_crawler import gsheet

from menu_crawler import cli
from menu_crawler import workbook as diag
from test_boards import book_with_domains
from test_foundation import launch_options
from test_modes import site  # noqa: F401 - pytest fixture


def scripted(monkeypatch, answers):
    it = iter(list(answers) + [""] * 6)          # a few spare Enters, then the input closes
    asked = []

    def fake(prompt=""):
        asked.append(prompt)
        try:
            return next(it)
        except StopIteration:
            raise EOFError from None
    monkeypatch.setattr("builtins.input", fake)
    return asked


def run(monkeypatch, tmp_path, answers):
    if launch_options():
        monkeypatch.setenv("MENU_CRAWLER_BROWSER_PATH", launch_options()["executable_path"])
    monkeypatch.chdir(tmp_path)
    asked = scripted(monkeypatch, answers)
    cli.main(["--browser", "chromium", "--headless", "--wait-ms", "0", "--delay-ms", "0"])
    return asked


def test_numbers_dedupe_crawl_fill_and_rerun_adds_nothing(site, tmp_path, monkeypatch):  # noqa: F811
    book = book_with_domains(tmp_path / "작업.xlsx")
    wb = load_workbook(book)
    wb["02_빈시트"]["D6"] = site                 # 02 -> local site; 01 keeps www.example.com (never crawled)
    wb.save(book)
    before = book.read_bytes()
    run(monkeypatch, tmp_path, [str(book), "2, 02", "", "2", "n", "", ""])
    out = next(tmp_path.glob("작업_메뉴입력_*[0-9].xlsx"))
    assert book.read_bytes() == before
    ws = load_workbook(out)["02_빈시트"]
    first = [ws.cell(13, c).value for c in range(8, 12)]
    assert "고객센터" in first
    log = next(tmp_path.glob("결과/작업기록_*.tsv")).read_text(encoding="utf-8-sig").splitlines()
    assert log[0].startswith("진단 시트\t점검 주소") and log[1].split("\t")[4] == "입력 완료"
    # additional investigation on the filled copy: nothing new, no new file
    for f in tmp_path.glob("작업_메뉴입력_*"):
        os.utime(f)
    run(monkeypatch, tmp_path, [str(out), "02", "", "2", "n", "", ""])
    assert len(list(tmp_path.glob("작업_메뉴입력_*[0-9].xlsx"))) == 1
    assert len(list(tmp_path.glob("작업_메뉴입력_*_메뉴입력_*.xlsx"))) == 0


def test_unreachable_target_is_reported_as_failure(tmp_path, monkeypatch, capsys):
    book = book_with_domains(tmp_path / "작업.xlsx")
    wb = load_workbook(book)
    wb["02_빈시트"]["D6"] = "http://127.0.0.1:9/"     # nothing listens on the discard port
    wb.save(book)
    run(monkeypatch, tmp_path, [str(book), "02", "", "2", "n", "", ""])
    out = capsys.readouterr().out
    assert "수집 실패" in out and not list(tmp_path.glob("작업_메뉴입력_*[0-9].xlsx"))


def test_tokens_and_exact_names():
    assert cli.split_tokens("1 2, 29 - 34, 사이트 A") == ["1", "2", "29-34", "사이트 A"]
    items = [diag.Target("03_y", "m.example.com 모바일", d6="https://m.example.com"),
             diag.Target("04_z", "z", d6="https://www.example.com")]
    assert cli.exact_sheet("example.com", items) is None          # an address, not a sheet name
    assert [t.sheet for t in diag.match_url("https://example.com", items)] == ["04_z"]


def test_result_file_name_keeps_dots(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    data = {"config": {"site_name": "www.example.com", "url": "https://www.example.com/", "mode": "public",
                       "extra_hosts": []}, "menus": [], "issues": [], "pages": [], "state": "complete",
            "visited_count": 0, "limit_reached": False, "started_at": "", "finished_at": "", "phases_done": []}
    path = cli.save(data, "")
    assert path.name.startswith("www.example.com_메뉴목록_") and path.suffix == ".xlsx"


def test_last_work_file_is_offered_and_names_do_not_stack(tmp_path, monkeypatch):
    from menu_crawler import state
    from menu_crawler.boards import ExcelBoard
    monkeypatch.delenv("MENU_CRAWLER_NO_STATE")
    monkeypatch.setattr(state, "FILE", tmp_path / "last_run.json")
    book = book_with_domains(tmp_path / "작업_메뉴입력_20261112_080109.xlsx")
    assert state.last_workbook() == ""
    state.remember_workbook(book)
    assert state.last_workbook() == str(book.resolve())
    out = ExcelBoard(book).out.name
    assert out.startswith("작업_메뉴입력_") and out.count("_메뉴입력_") == 1


LEGACY = """<!doctype html><html lang="en"><head><title>Global Banking</title>
<script>function goMenu(c){ location.href='/menu.do?c=' + c; } function goForgot(){ location.href='/forgotid.do'; }
function goLogin(){ location.href='/'; } function closePop(){}</script></head><body>
<div id="wrap"><div id="topMenu"><ul>
 <li><a href="javascript:goLogin()">Login</a></li>
 <li><a href="javascript:goMenu('A01')">Inquiry</a><ul>
   <li><a href="javascript:goMenu('A0101')">Account Inquiry</a><ul><li><a href="javascript:goMenu('A010101')">Balance</a></li>
     <li><a href="javascript:goMenu('A010102')">Transaction History</a></li></ul></li></ul></li>
 <li><a href="javascript:goMenu('B01')">Transfer</a><ul><li><a href="javascript:goMenu('B0101')">Domestic Transfer</a></li></ul></li>
 <li><a href="javascript:fnLogout()">Logout</a></li>
 <li><a href="javascript:void(0)" class="btn-all">Open Whole Menu</a></li>
</ul></div>
<div id="allMenu"><ul><li><a href="javascript:goMenu('A0101')">Account Inquiry</a><ul><li><a href="javascript:goMenu('A010101')">Balance</a></li></ul></li></ul></div>
<div class="login-box"><form action="/login" method="post"><input name="uid"><input type="password" name="pw">
 <button type="submit">Login</button></form>
 <a href="javascript:goForgot()">Forgot User ID</a> <a href="/forgotpw.do">Forgot Password</a></div>
<div class="layer_popup" style="position:fixed;left:0;top:0;width:600px;height:500px;z-index:999;background:#fff">
 Notice <a href="javascript:closePop()">Do not show today</a></div>
</div><script>window.open('/event', 'ev');</script></body></html>"""
FORGOT_ID = """<html lang="en"><body><h1>Forgot User ID</h1><ul class="tab"><li><a href="#t1">User ID issued at the branch</a></li>
<li><a href="#t2">User ID registered online</a></li></ul></body></html>"""


def test_script_menus_on_a_legacy_english_site_are_filled(tmp_path, monkeypatch):
    import threading
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

    class Legacy(BaseHTTPRequestHandler):
        def do_GET(self):
            pages = {"/": LEGACY, "/index.do": LEGACY, "/forgotid.do": FORGOT_ID}
            body = pages.get(self.path.split("?")[0], "<html><body>page</body></html>").encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(body)

        def log_message(self, *a):
            pass
    server = ThreadingHTTPServer(("127.0.0.1", 0), Legacy)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        book = book_with_domains(tmp_path / "작업.xlsx")
        wb = load_workbook(book)
        wb["02_빈시트"]["D6"] = f"http://127.0.0.1:{server.server_address[1]}/"
        wb.save(book)
        run(monkeypatch, tmp_path, [str(book), "02", "", "2", "n", "", "", ""])
        out = next(tmp_path.glob("작업_메뉴입력_*[0-9].xlsx"))
        ws = load_workbook(out)["02_빈시트"]
        cols = [[ws.cell(r, c).value for r in (13, 14, 15)] for c in range(8, 30) if ws.cell(13, c).value]
        assert cols[:3] == [["Login", "Forgot User ID", "User ID issued at the branch"],
                            ["Login", "Forgot User ID", "User ID registered online"], ["Login", "Forgot Password", None]]
        assert ["Inquiry", "Account Inquiry", "Balance"] in cols and ["Transfer", "Domestic Transfer", None] in cols
        assert ["Account Inquiry", "Balance", None] not in cols                     # copy from the whole-menu panel
        names = [x for col in cols for x in col if x]
        assert not {"Logout", "Open Whole Menu", "Do not show today"} & set(names)
        assert ["Login", None, None] not in cols and ["Inquiry", None, None] not in cols
        # the menu-list files already on disk can fill a fresh copy without crawling again (two files: 2nd = duplicates)
        import shutil
        first_result = next((tmp_path / "결과").glob("*_메뉴목록_*.xlsx"))
        shutil.copy(first_result, first_result.with_name("사본_메뉴목록_2.xlsx"))
        cli.main(["--from-result", str(tmp_path / "결과"), "--workbook", str(book), "--workbook-out", str(tmp_path / "re.xlsx"),
                  "--include-unverified", "-y"])
        ws2 = load_workbook(tmp_path / "re.xlsx")["02_빈시트"]
        again = [[ws2.cell(r, c).value for r in (13, 14, 15)] for c in range(8, 14) if ws2.cell(13, c).value]
        assert again == cols
    finally:
        server.shutdown()


BOOK = "1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789"
FORM = {"A13": "1 Depth · 최상위 (필수)", "A14": "2 Depth · 하위 (선택)", "A15": "3 Depth · 세부 (선택)"}


@pytest.fixture()
def gas(tmp_path, monkeypatch):
    """apps_script/MenuImport.gs running unchanged in a Google Apps Script stand-in (node)."""
    if shutil.which("node") is None:
        pytest.skip("node not installed")
    def make(sheets):
        fixture = tmp_path / "gas.json"
        fixture.write_text(json.dumps({"token": "k", "books": {BOOK: {"name": "진단표", "sheets": sheets}}}, ensure_ascii=False))
        proc = subprocess.Popen(["node", str(Path(__file__).parent / "js" / "gas_server.js"), str(fixture)],
                                stdout=subprocess.PIPE, text=True)
        port = int(proc.stdout.readline().split()[1])
        procs.append(proc)
        monkeypatch.setenv("MENU_CRAWLER_GSHEET_TEST", "1")
        monkeypatch.setattr(gsheet, "load_config", lambda: {"url": f"http://127.0.0.1:{port}/", "token": "k"})
        state = lambda: json.loads(urllib.request.urlopen(f"http://127.0.0.1:{port}/__state").read())[BOOK]  # noqa: E731
        return port, state
    procs = []
    yield make
    for proc in procs:
        proc.kill()


def test_google_sheet_full_flow_and_rerun(site, tmp_path, monkeypatch, gas):  # noqa: F811
    _, state = gas([
        {"name": "00_Sample", "cells": {"H13": "예시"}},
        {"name": "05_사이트", "cells": {**FORM, "D3": "사이트", "D6": site, "H13": "기존 메뉴", "H20": "양호"},
         "formulas": {"H16": "=MID(H13,1,9)"}},
        {"name": "07_숨김", "hidden": True, "cells": {"D6": site}},
    ])
    url = f"https://docs.google.com/spreadsheets/d/{BOOK}/edit#gid=0"
    run(monkeypatch, tmp_path, ["", url, "05", "", "2", "n", "", "", ""])
    sheet = {s["name"]: s for s in state()}["05_사이트"]
    cells = sheet["cells"]
    assert cells["H13"] == "기존 메뉴" and cells["H20"] == "양호"                  # existing data untouched
    assert "고객센터" in [cells.get(f"{c}13") for c in "IJKLM"]
    assert all(k[1:] in {"13", "14", "15"} for k in sheet["writes"])               # values only in menu rows
    assert not sheet["writes"][0].startswith("H")                                  # after the last used column
    assert {s["name"]: s for s in state()}["07_숨김"]["writes"] == []
    before = len(sheet["writes"])
    run(monkeypatch, tmp_path, [url, "05", "", "2", "n", "", "", ""])              # address pasted at the first prompt
    assert len({s["name"]: s for s in state()}["05_사이트"]["writes"]) == before     # nothing new
    assert not list(tmp_path.glob("*_메뉴입력_*.xlsx"))


def test_google_sheet_skips_merges_formulas_and_hidden_columns(tmp_path, gas):
    port, state = gas([{"name": "06_병합", "cells": {**FORM, "H13": "기존", "I14": "탈퇴"}, "merges": ["H13:I13"],
                        "formulas": {"J13": "=A1"}, "hiddenCols": ["K"]},
                       {"name": "08_다른양식", "cells": {"A12": "1 Depth", "A13": "2 Depth", "A14": "3 Depth"}}])
    client = gsheet.Client(f"http://127.0.0.1:{port}/", "k", spreadsheet=f"https://docs.google.com/spreadsheets/d/{BOOK}/edit")
    res = client.fill("06_병합", [["기존", "탈퇴"], ["새1"], ["새2"]])
    assert res["duplicates"] == [["기존", "탈퇴", ""]] and res["columns"] == [12, 13] and res["range"] == "L13:M15"
    assert res["url"].endswith("/edit#gid=1000")
    cells = {x["name"]: x for x in state()}["06_병합"]["cells"]
    assert (cells["L13"], cells["M13"]) == ("새1", "새2") and "K13" not in cells and "J13" not in cells
    with pytest.raises(gsheet.SheetError, match="양식"):
        client.fill("08_다른양식", [["새"]])                                          # another layout: nothing written
    assert {x["name"]: x for x in state()}["08_다른양식"]["writes"] == []
    with pytest.raises(gsheet.SheetError):
        gsheet.Client(f"http://127.0.0.1:{port}/", "wrong", spreadsheet=BOOK).ping()


def test_auto_mode_logs_in_only_when_the_site_shows_just_a_login_screen(site, tmp_path, monkeypatch):  # noqa: F811
    asked_login = []
    monkeypatch.setattr(cli, "login_step", lambda job, config, username=None: asked_login.append(config.url) or False)
    book = book_with_domains(tmp_path / "작업.xlsx")
    wb = load_workbook(book)
    wb["02_빈시트"]["D6"] = site                      # public menus visible -> no login step
    wb.save(book)
    run(monkeypatch, tmp_path, [str(book), "02", "", "1", "n", "", "", "", ""])
    assert asked_login == []


def test_needs_login_rules():
    def report(menus, pages=()):
        return {"menus": menus, "pages": [{"status": s} for s in pages]}
    def menu(path, evidence="메뉴 DOM", status="수집됨", kind="link"):
        return {"path": path, "evidence": evidence, "status": status, "kind": kind, "ready": status == "수집됨",
                "hierarchy_review": False, "id": "/".join(path), "url": "", "visibility": "", "depths": (path + ["", ""])[:3]}
    login_only = report([menu(["Login", "Forgot User ID"], "로그인 화면"), menu(["Login", "Forgot Password"], "로그인 화면")])
    assert cli.needs_login(login_only)
    public = report([menu(["Login", "Forgot User ID"], "로그인 화면")] + [menu([f"메뉴{i}"]) for i in range(5)], ["수집됨"] * 6)
    assert not cli.needs_login(public)
    # every menu leads to a login wall: names are known but nothing is confirmed
    walls = report([menu([f"메뉴{i}"], status="로그인 필요") for i in range(6)], ["수집됨"] + ["로그인 필요"] * 7)
    assert cli.needs_login(walls)
    assert not cli.needs_login(report([], ["접속 실패"]))           # unreachable: logging in would not help
    assert cli.needs_login(report([]))


def test_auto_mode_goes_to_login_for_a_login_only_site(tmp_path, monkeypatch):
    import threading
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
    page = ("<html lang='en'><body><div class='login'><form method='post' action='/login'><input name='id'>"
            "<input type='password' name='pw'><button type='submit'>Login</button></form>"
            "<a href='/find-id'>Forgot User ID</a> <a href='/find-pw'>Forgot Password</a></div></body></html>")

    class LoginOnly(BaseHTTPRequestHandler):
        def do_GET(self):
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(page.encode())

        def log_message(self, *a):
            pass
    server = ThreadingHTTPServer(("127.0.0.1", 0), LoginOnly)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    asked = []
    monkeypatch.setattr(cli, "login_step", lambda job, config, username=None: asked.append(config.url) or False)
    try:
        book = book_with_domains(tmp_path / "작업.xlsx")
        wb = load_workbook(book)
        wb["02_빈시트"]["D6"] = f"http://127.0.0.1:{server.server_address[1]}/"
        wb.save(book)
        run(monkeypatch, tmp_path, [str(book), "02", "", "1", "n", "", "", "", ""])
        assert len(asked) == 1                       # the login step was offered for this target
        ws = load_workbook(next(tmp_path.glob("작업_메뉴입력_*[0-9].xlsx")))["02_빈시트"]
        assert [ws["H13"].value, ws["H14"].value] == ["Login", "Forgot User ID"]       # login functions first
    finally:
        server.shutdown()


def test_my_browser_mode_attaches_crawls_and_leaves_the_browser_open(site, tmp_path, monkeypatch):  # noqa: F811
    from menu_crawler import mybrowser
    from playwright.sync_api import sync_playwright
    if not launch_options():
        pytest.skip("needs MENU_CRAWLER_BROWSER_PATH")
    port = 9300 + os.getpid() % 500
    monkeypatch.setattr(mybrowser, "PORT", port)
    monkeypatch.setattr(mybrowser, "PROFILE", tmp_path / "browser_profile")
    book = book_with_domains(tmp_path / "작업.xlsx")
    wb = load_workbook(book)
    wb["02_빈시트"]["D6"] = site
    wb.save(book)
    try:
        mybrowser.start(site + "/notice", headless=True)          # the user's own window with a tab of theirs
        if launch_options():
            monkeypatch.setenv("MENU_CRAWLER_BROWSER_PATH", launch_options()["executable_path"])
        monkeypatch.chdir(tmp_path)
        scripted(monkeypatch, [str(book), "02", "", "2", "n", "", "", "", ""])
        cli.main(["--my-browser", "--headless", "--wait-ms", "0", "--delay-ms", "0"])
        ws = load_workbook(next(tmp_path.glob("작업_메뉴입력_*[0-9].xlsx")))["02_빈시트"]
        assert "고객센터" in [ws.cell(13, c).value for c in range(8, 14)]
        assert mybrowser.running(port)                             # the browser is still open
        with sync_playwright() as pw:
            urls = [p.url for p in pw.chromium.connect_over_cdp(mybrowser.endpoint(port)).contexts[0].pages]
        assert any(u.endswith("/notice") for u in urls)           # the user's tab is untouched, ours is closed
        assert not any(u.rstrip("/") == site for u in urls)
    finally:
        subprocess.run(["pkill", "-f", str(tmp_path / "browser_profile")], check=False)


def test_login_step_can_fall_back_to_public_menus(site, tmp_path, monkeypatch):  # noqa: F811
    monkeypatch.setattr(cli, "login_step", lambda job, config, username=None: "public")
    book = book_with_domains(tmp_path / "작업.xlsx")
    wb = load_workbook(book)
    wb["02_빈시트"]["D6"] = site
    wb.save(book)
    # mode 3 = after login only; the account fails, so 'p' collects the public menus instead
    run(monkeypatch, tmp_path, [str(book), "02", "", "3", "1", "n", "n", "", "", "", ""])
    ws = load_workbook(next(tmp_path.glob("작업_메뉴입력_*[0-9].xlsx")))["02_빈시트"]
    assert "고객센터" in [ws.cell(13, c).value for c in range(8, 14)]


def test_failed_targets_can_be_retried_in_the_same_run(tmp_path, monkeypatch):
    book = book_with_domains(tmp_path / "작업.xlsx")
    wb = load_workbook(book)
    wb["02_빈시트"]["D6"] = "http://127.0.0.1:9/"            # nothing listens: the crawl fails
    wb.save(book)
    asked = run(monkeypatch, tmp_path, [str(book), "02", "", "2", "n", "", "y", "n"])
    retries = [a for a in asked if "다시 시도할까요" in a]
    assert len(retries) == 2                                  # offered, retried, offered again, declined


def test_session_loss_after_login_is_detected():
    import threading
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
    from menu_crawler.crawler import Crawler
    from menu_crawler.models import CrawlConfig
    login = "<form><input name=id><input type=password name=pw><button type=submit>Login</button></form>"
    home = "<nav><ul>" + "".join(f"<li><a href='/m{i}'>메뉴{i}</a></li>" for i in range(6)) + "</ul></nav>"

    class Expired(BaseHTTPRequestHandler):
        def do_GET(self):
            body = home if self.path == "/" else "<html><body>세션이 만료되었습니다." + login + "</body></html>"
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(body.encode())

        def log_message(self, *a):
            pass
    server = ThreadingHTTPServer(("127.0.0.1", 0), Expired)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        config = CrawlConfig(url=f"http://127.0.0.1:{server.server_address[1]}/", mode="login", headless=True,
                             browser="chromium", wait_ms=0, delay_ms=0, click_menus=False)
        engine = Crawler(config)
        engine.open()
        try:
            engine.crawl("after")
            report = engine.snapshot()
        finally:
            engine.close()
        assert report["session_lost"] and report["visited_count"] <= 5
    finally:
        server.shutdown()


def test_last_targets_and_mode_are_reused_with_enter(site, tmp_path, monkeypatch):  # noqa: F811
    from menu_crawler import state
    monkeypatch.setattr(state, "FILE", tmp_path / "last_run.json")
    book = book_with_domains(tmp_path / "작업.xlsx")
    wb = load_workbook(book)
    wb["02_빈시트"]["D6"] = site
    wb.save(book)
    monkeypatch.delenv("MENU_CRAWLER_NO_STATE")
    run(monkeypatch, tmp_path, [str(book), "02", "", "2", "n", "", ""])
    assert state.pref("targets") == "02" and state.pref("mode") == "public"
    asked = run(monkeypatch, tmp_path, [str(book), "", "", "", "n", "", ""])     # Enter = last targets, last mode
    assert any("지난번 02" in a for a in asked) and any("번호 [2]" in a for a in asked)


def test_cleanup_check_on_google_sheet(tmp_path, monkeypatch, gas):
    port, state_ = gas([{"name": "09_정리", "cells": {
        **FORM, "D6": "https://a.example.com",
        "H13": "Open/Close Menu",
        "I13": "Customer Center",
        "J13": "Customer Center", "J14": "User Guide", "J15": "Demo",
        "K13": "User Guide", "K14": "Demo",
        "L13": "FAQ", "M13": "Customer Center", "M14": "FAQ", "L20": "양호"}}])
    url = f"https://docs.google.com/spreadsheets/d/{BOOK}/edit"
    run(monkeypatch, tmp_path, ["c", url, "", "y"])
    cells = state_()[0]["cells"]
    assert "H13" not in cells and "I13" not in cells and "K13" not in cells and "K14" not in cells   # cleared
    assert cells["J13"] == "Customer Center" and cells["M14"] == "FAQ"                              # real menus stay
    assert cells["L13"] == "FAQ" and cells["L20"] == "양호"                                          # has a result: kept
    assert all(k[1:] in {"13", "14", "15"} for k in state_()[0]["writes"])


def test_cleanup_check_on_excel_copy(tmp_path, monkeypatch):
    book = book_with_domains(tmp_path / "작업.xlsx")
    wb = load_workbook(book)
    ws = wb["02_빈시트"]
    ws["H13"], ws["I13"], ws["I14"], ws["J13"] = "전체메뉴 열기/닫기", "조회", "계좌조회", "조회"
    wb.save(book)
    before = book.read_bytes()
    run(monkeypatch, tmp_path, ["c", str(book), "02", "y"])
    out = next(tmp_path.glob("작업_메뉴입력_*[0-9].xlsx"))
    ws = load_workbook(out)["02_빈시트"]
    assert (ws["H13"].value, ws["J13"].value) == (None, None) and (ws["I13"].value, ws["I14"].value) == ("조회", "계좌조회")
    assert book.read_bytes() == before


def test_saved_account_is_only_typed_into_the_same_site(monkeypatch):
    filled = []

    class FakeJob:
        def command(self, name, **kw):
            filled.append((name, kw.get("username"), kw.get("password")))
            return 1

        def wait_done(self, ticket):
            return {"message": ""}
    monkeypatch.setattr(cli, "run_command", lambda job, name, **kw: {"message": ""})
    monkeypatch.setattr(cli.Account, "username", "consultant")
    monkeypatch.setattr(cli.Account, "password", "pw")
    monkeypatch.setattr(cli.Account, "site", "client-a.com")
    from menu_crawler.models import CrawlConfig
    scripted(monkeypatch, [""])
    cli.login_step(FakeJob(), CrawlConfig(url="https://www.client-b.com/"))
    assert filled == []                                          # another client's site: nothing typed
    scripted(monkeypatch, [""])
    cli.login_step(FakeJob(), CrawlConfig(url="https://dev.client-a.com/"))
    assert filled == [("fill", "consultant", "pw")]                 # same site: filled
    assert cli.site_of("https://u2lusdev.example.co.kr/x") == "example.co.kr"



def test_account_pin_field_is_not_a_login_screen():
    import threading
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
    from menu_crawler.crawler import Crawler
    from menu_crawler.models import CrawlConfig
    nav = "<nav><ul>" + "".join(f"<li><a href='/t{i}'>이체{i}</a></li>" for i in range(5)) + "</ul></nav>"
    pin = "<form><input type=password name=pin placeholder='계좌 비밀번호'><button type=submit>이체</button></form>"

    class Bank(BaseHTTPRequestHandler):
        def do_GET(self):
            body = f"<html><body>{nav}{'' if self.path == '/' else pin}</body></html>"
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(body.encode())

        def log_message(self, *a):
            pass
    server = ThreadingHTTPServer(("127.0.0.1", 0), Bank)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        engine = Crawler(CrawlConfig(url=f"http://127.0.0.1:{server.server_address[1]}/", mode="login", headless=True,
                                     browser="chromium", wait_ms=0, delay_ms=0, click_menus=False))
        engine.open()
        try:
            engine.crawl("after")
            report = engine.snapshot()
        finally:
            engine.close()
        assert not report["session_lost"]
        assert all(p["status"] == "수집됨" for p in report["pages"])
        assert not any(m["evidence"] == "로그인 화면" for m in report["menus"])
    finally:
        server.shutdown()


TRAPS = """<!doctype html><html><head><script>
function go(p){ location.href = p; }
function doLogout(){ fetch('/hit/logout', {method:'POST'}); }
function runTransfer(){ fetch('/api/transfer', {method:'POST', body:'amount=1000000'}); }
function post(p){ fetch(p, {method:'POST'}); }
</script></head><body><nav><ul>
 <li><a href="javascript:go('/notice')">공지사항</a></li>
 <li><a href="javascript:go('/rates')">Exchange rate</a></li>
 <li><a href="javascript:runTransfer()">이체 실행</a></li>
 <li><a href="javascript:post('/hit/submit-en')">Submit transfer</a></li>
 <li><a href="javascript:doLogout()">마이페이지</a></li>
 <li><a href="javascript:post('/hit/withdraw')" title="내 정보">회원탈퇴</a></li>
 <li><a href="javascript:post('/hit/save')">Save changes</a></li>
 <li><form action="/api/form-submit" method="post"><input type=hidden name=amount value=1000000>
     <button type=button onclick="this.form.submit()">자동 처리</button></form></li>
 <li><span role=menuitem onclick="document.forms[0].requestSubmit()">빠른 처리</span></li>
</ul></nav></body></html>"""


def test_crawler_never_triggers_state_changing_requests():
    import threading
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
    from menu_crawler.crawler import Crawler
    from menu_crawler.models import CrawlConfig
    seen = []

    class Traps(BaseHTTPRequestHandler):
        def respond(self):
            seen.append((self.command, self.path.split("?")[0]))
            body = TRAPS if self.path == "/" else "<html><body><ul class='tab'><li><a href='#a'>탭A</a></li></ul></body></html>"
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(body.encode())

        do_GET = do_POST = respond

        def log_message(self, *a):
            pass
    server = ThreadingHTTPServer(("127.0.0.1", 0), Traps)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        engine = Crawler(CrawlConfig(url=f"http://127.0.0.1:{server.server_address[1]}/", headless=True,
                                     browser="chromium", wait_ms=0, delay_ms=0))
        engine.open()
        try:
            engine.crawl("before")
            report = engine.snapshot()
        finally:
            engine.close()
        posts = [p for m, p in seen if m != "GET"]
        assert posts == [], posts                                              # nothing state-changing was sent
        clicked = {tuple(m["path"]) for m in report["menus"] if "눌러" in (m.get("reason") or "") and m["status"] == "수집됨"}
        assert ("공지사항",) in clicked and ("Exchange rate",) in clicked        # navigation menus were opened
    finally:
        server.shutdown()


def test_google_writer_guards(tmp_path, monkeypatch, gas):
    port, state_ = gas([{"name": "00_Sample", "cells": {**FORM}},
                        {"name": "05_숨김", "hidden": True, "cells": {**FORM}},
                        {"name": "06_병합정리", "cells": {**FORM, "H13": "고객센터", "I14": "공지", "J14": "FAQ"}, "merges": ["H13:J13"]},
                        {"name": "07_다른양식", "cells": {"A13": "11 Depth", "A14": "2 Depth", "A15": "3 Depth"}}])
    client = gsheet.Client(f"http://127.0.0.1:{port}/", "k", spreadsheet=BOOK)
    for name in ("00_Sample", "05_숨김", "07_다른양식"):
        with pytest.raises(gsheet.SheetError):
            client.fill(name, [["x"]])
    res = client.clear("06_병합정리", [{"col": 8, "menu": ["고객센터", "", ""]}])
    assert res["cleared"] == [] and res["skipped"] == [8]                # a merged label is never emptied
    assert all(not s["writes"] for s in state_())


def test_google_addresses_and_test_hatch(monkeypatch):
    sid = "1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789"
    assert gsheet.sheet_ref(f"https://docs.google.com/spreadsheets/u/1/d/{sid}/edit#gid=5").endswith(f"/d/{sid}/edit")
    assert gsheet.sheet_ref(f"docs.google.com/spreadsheets/d/{sid}/edit") != ""
    monkeypatch.setenv("MENU_CRAWLER_GSHEET_TEST", "0")
    with pytest.raises(gsheet.SheetError):
        gsheet.Client("http://127.0.0.1:1/", "k")
    monkeypatch.setenv("MENU_CRAWLER_GSHEET_TEST", "1")
    with pytest.raises(gsheet.SheetError):
        gsheet.Client("http://127.0.0.1:1@evil.example/x", "k")


def test_wrong_layout_sheet_is_refused_before_crawling(tmp_path, monkeypatch, capsys):
    book = book_with_domains(tmp_path / "작업.xlsx")
    wb = load_workbook(book)
    ws = wb["02_빈시트"]
    ws["A13"], ws["A14"], ws["A15"], ws["D6"] = "메뉴 번호", "1 Depth", "2 Depth", "https://x.example.com"
    wb.save(book)
    with pytest.raises(SystemExit):                     # nothing valid to crawl: it keeps asking for a target
        run(monkeypatch, tmp_path, [str(book), "02"])
    out = capsys.readouterr().out
    assert "기입할 수 없습니다" in out and "수집 방식을 선택하세요" not in out


def test_outdated_apps_script_deployment_is_reported(tmp_path, monkeypatch, gas, capsys):
    import re as _re
    source = (Path(__file__).parent.parent / "apps_script" / "MenuImport.gs").read_text(encoding="utf-8")
    assert int(_re.search(r"var VERSION = (\d+);", source).group(1)) == gsheet.SCRIPT_VERSION      # kept in step
    port, _ = gas([{"name": "05_사이트", "cells": {**FORM}}])
    client = gsheet.Client(f"http://127.0.0.1:{port}/", "k", spreadsheet=BOOK)
    client.targets()
    assert not client.outdated()
    monkeypatch.setattr(gsheet, "SCRIPT_VERSION", gsheet.SCRIPT_VERSION + 1)                        # deployed copy is older
    assert client.outdated()
    with pytest.raises(SystemExit):                                                                  # default answer: stop
        run(monkeypatch, tmp_path, [f"https://docs.google.com/spreadsheets/d/{BOOK}/edit", ""])
    assert "예전 버전" in capsys.readouterr().out


def test_menus_can_be_left_out_before_the_google_fill(site, tmp_path, monkeypatch, gas, capsys):  # noqa: F811
    from menu_crawler import ui
    _, state_ = gas([{"name": "05_사이트", "cells": {**FORM, "D6": site}}])
    url = f"https://docs.google.com/spreadsheets/d/{BOOK}/edit"
    run(monkeypatch, tmp_path, [url, "05", "", "2", "n", "", "", "v", "1", ""])     # v = list, leave out number 1, Enter = fill
    cells = state_()[0]["cells"]
    names = [cells.get(f"{c}13") for c in "HIJK"]
    assert "로그인" not in names and "고객센터" in names and names[2:] == [None, None]   # two of three written, no gap
    assert "골라서 뺀 메뉴 1개" in capsys.readouterr().out
    assert ui.pick_numbers("3, 7-9 12", 12) == {3, 7, 8, 9, 12} and ui.pick_numbers("", 5) == set()
    with pytest.raises(ValueError):
        ui.pick_numbers("0", 5)
    with pytest.raises(ValueError):
        ui.pick_numbers("a", 5)
