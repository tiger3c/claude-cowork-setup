"""The whole interactive flow: work Excel -> targets -> crawl -> fill, driven by scripted answers."""
import os

from openpyxl import load_workbook

from menu_crawler import cli
from menu_crawler import workbook as diag
from test_boards import book_with_domains
from test_foundation import launch_options
from test_modes import site  # noqa: F401 - pytest fixture


def scripted(monkeypatch, answers):
    it = iter(answers)
    asked = []

    def fake(prompt=""):
        asked.append(prompt)
        return next(it)
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
    run(monkeypatch, tmp_path, [str(book), "2, 02", "", "1", "n", "", ""])
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
    run(monkeypatch, tmp_path, [str(out), "02", "", "1", "n", "", ""])
    assert len(list(tmp_path.glob("작업_메뉴입력_*[0-9].xlsx"))) == 1
    assert len(list(tmp_path.glob("작업_메뉴입력_*_메뉴입력_*.xlsx"))) == 0


def test_unreachable_target_is_reported_as_failure(tmp_path, monkeypatch, capsys):
    book = book_with_domains(tmp_path / "작업.xlsx")
    wb = load_workbook(book)
    wb["02_빈시트"]["D6"] = "http://127.0.0.1:9/"     # nothing listens on the discard port
    wb.save(book)
    run(monkeypatch, tmp_path, [str(book), "02", "", "1", "n", "", ""])
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


LEGACY = """<!doctype html><html lang="en"><head><title>Global Banking</title></head><body>
<div id="wrap"><div id="topMenu"><ul>
 <li><a href="javascript:goMenu('A01')">Inquiry</a><ul>
   <li><a href="javascript:goMenu('A0101')">Account Inquiry</a><ul><li><a href="javascript:goMenu('A010101')">Balance</a></li>
     <li><a href="javascript:goMenu('A010102')">Transaction History</a></li></ul></li></ul></li>
 <li><a href="javascript:goMenu('B01')">Transfer</a><ul><li><a href="javascript:goMenu('B0101')">Domestic Transfer</a></li></ul></li>
 <li><a href="javascript:fnLogout()">Logout</a></li>
</ul></div>
<div class="layer_popup" style="position:fixed;left:0;top:0;width:600px;height:500px;z-index:999;background:#fff">
 Notice <a href="javascript:closePop()">Do not show today</a></div>
</div><script>window.open('/event', 'ev');</script></body></html>"""


def test_script_menus_on_a_legacy_english_site_are_filled(tmp_path, monkeypatch):
    import threading
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

    class Legacy(BaseHTTPRequestHandler):
        def do_GET(self):
            body = LEGACY.encode() if self.path in ("/", "/index.do") else b"<html><body>event</body></html>"
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
        asked = run(monkeypatch, tmp_path, [str(book), "02", "", "1", "n", "", "", ""])
        assert any("주소 없는 메뉴" in a or "넣을까요" in a for a in asked)
        out = next(tmp_path.glob("작업_메뉴입력_*[0-9].xlsx"))
        ws = load_workbook(out)["02_빈시트"]
        cols = [[ws.cell(r, c).value for r in (13, 14, 15)] for c in range(8, 14) if ws.cell(13, c).value]
        assert ["Inquiry", "Account Inquiry", "Balance"] in cols
        assert ["Transfer", "Domestic Transfer", None] in cols
        assert not any("Logout" in str(x) for col in cols for x in col)
        assert not any(c[0] == "Inquiry" and c[1] is None for c in cols)     # parent is not its own column
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
