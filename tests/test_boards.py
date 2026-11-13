"""진단 대상 찾기(대표 도메인 매핑), 여러 대상 차례 기입, 구글 시트 연결."""
import io
import json
import shutil
import subprocess
import zipfile
from pathlib import Path

import pytest
from openpyxl import load_workbook

from menu_crawler import gsheet
from menu_crawler import workbook as diag
from menu_crawler.boards import ExcelBoard, SheetsBoard
from test_workbook import diagnosis_book


def book_with_domains(path):
    diagnosis_book(path)
    wb = load_workbook(path)
    wb["01_테스트"]["D3"], wb["01_테스트"]["D6"] = "테스트 사이트", "www.example.com/bank/index.jsp"
    wb["02_빈시트"]["D6"] = "예) https://www.example.com"            # placeholder, not a domain
    ws = wb.copy_worksheet(wb["02_빈시트"])
    ws.title = "03_다른사이트"
    ws["D6"] = "https://other.example.org"
    wb.save(path)
    return path


def test_targets_read_name_domain_and_next_cell(tmp_path):
    items = {t.sheet: t for t in diag.targets(book_with_domains(tmp_path / "d.xlsx"))}
    assert list(items) == ["01_테스트", "02_빈시트", "03_다른사이트"]
    assert (items["01_테스트"].name, items["01_테스트"].domain) == ("테스트 사이트", "https://www.example.com/bank/index.jsp")
    assert items["01_테스트"].menus == 1 and items["01_테스트"].cell == "J13"      # H menu, I result -> J
    assert items["02_빈시트"].domain == "" and items["02_빈시트"].cell == "H13"


def test_url_maps_to_sheet_by_domain(tmp_path):
    items = diag.targets(book_with_domains(tmp_path / "d.xlsx"))
    assert [t.sheet for t in diag.match_url("https://example.com/bank/main.do", items)] == ["01_테스트"]
    assert [t.sheet for t in diag.match_url("other.example.org/login", items)] == ["03_다른사이트"]
    assert diag.match_url("https://www.example.com/card/", items) == []          # same host, other service path
    assert diag.match_url("https://unknown.test/", items) == []


def test_select_targets_by_name(tmp_path):
    items = diag.targets(book_with_domains(tmp_path / "d.xlsx"))
    assert [t.sheet for t in diag.select_targets("빈시트, 다른", items)] == ["02_빈시트", "03_다른사이트"]
    with pytest.raises(ValueError):
        diag.select_targets("사이트", items)         # ambiguous: 테스트 사이트, 다른사이트
    with pytest.raises(ValueError):
        diag.select_targets("없는대상", items)


def test_excel_board_fills_several_targets_into_one_new_file(tmp_path):
    src = book_with_domains(tmp_path / "작업.xlsx")
    before = src.read_bytes()
    board = ExcelBoard(src, tmp_path / "작업_메뉴입력.xlsx")
    items = {t.sheet: t for t in board.targets()}
    a = board.fill(items["01_테스트"], [["조회", "계좌"], ["기존", "하위"]], False, tmp_path)
    b = board.fill(items["03_다른사이트"], [["공지"]], False, tmp_path)
    assert (a.status, a.where, a.written, a.duplicates) == ("입력 완료", "J13:J15", 1, 1)
    assert (b.status, b.where) == ("입력 완료", "H13:H15")
    assert src.read_bytes() == before                              # original untouched
    wb = load_workbook(board.out)
    assert (wb["01_테스트"]["H13"].value, wb["01_테스트"]["J13"].value, wb["01_테스트"]["J14"].value) == ("기존", "조회", "계좌")
    assert wb["03_다른사이트"]["H13"].value == "공지"
    with zipfile.ZipFile(board.out) as z:
        assert any(n.startswith("xl/charts/") for n in z.namelist())




class FakeResponse(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def test_sheets_board_uses_web_app(monkeypatch, tmp_path):
    calls = []

    def fake_urlopen(req, timeout):
        body = json.loads(req.data)
        calls.append(body)
        assert body["token"] == "k"
        if body["action"] == "targets":
            data = {"ok": True, "title": "진단표", "targets": [{"sheet": "29_대상", "name": "대상", "domain": "a.example.com", "menus": 2, "start": 10}]}
        else:
            write = [m for m in body["menus"] if m != ["기존", "", ""]]
            data = {"ok": True, "start": 10, "cell": "J13", "range": "J13:J15", "existing": 2, "write": write,
                    "duplicates": [["기존", "", ""]], "overflow": [], "written": 0 if body["dryRun"] else len(write)}
        return FakeResponse(json.dumps(data).encode())

    monkeypatch.setattr(gsheet.urllib.request, "urlopen", fake_urlopen)
    board = SheetsBoard(gsheet.Client("https://script.google.com/macros/s/x/exec", "k"))
    [target] = board.targets()
    assert (target.sheet, target.domain, target.cell) == ("29_대상", "https://a.example.com", "J13")
    res = board.fill(target, [["기존"], ["새 메뉴", "하위\t메뉴"]], False, tmp_path)
    assert (res.status, res.where, res.written, res.duplicates) == ("입력 완료", "J13:J15", 1, 1)
    assert [c["dryRun"] for c in calls[1:]] == [True, False]
    assert calls[2]["menus"] == [["기존", "", ""], ["새 메뉴", "하위 메뉴", ""]]       # cleaned before sending


def test_sheets_client_reports_errors(monkeypatch):
    monkeypatch.setattr(gsheet.urllib.request, "urlopen", lambda req, timeout: FakeResponse(b'{"ok": false, "error": "key"}'))
    with pytest.raises(gsheet.SheetError, match="key"):
        gsheet.Client("https://script.google.com/macros/s/x/exec", "bad").ping()
    with pytest.raises(gsheet.SheetError):
        gsheet.Client("https://evil.example/exec", "k")


@pytest.mark.skipif(shutil.which("node") is None, reason="node not installed")
def test_apps_script_planning_logic():
    script = Path(__file__).parent / "js" / "menu_import.test.js"
    out = subprocess.run(["node", str(script)], capture_output=True, text=True, timeout=60)
    assert out.returncode == 0, out.stderr


def test_several_addresses_in_one_d6_cell_and_hint_ignored(tmp_path):
    path = book_with_domains(tmp_path / "d.xlsx")
    wb = load_workbook(path)
    wb["02_빈시트"]["D6"] = "dev.example.net/app\nm.example.net, 예) https://www.example.com"
    wb.save(path)
    items = diag.targets(path)
    assert [t.sheet for t in diag.match_url("https://m.example.net/main", items)] == ["02_빈시트"]
    assert [t.sheet for t in diag.match_url("http://dev.example.net/app/login.do", items)] == ["02_빈시트"]
    assert diag.match_url("https://dev.example.net/other", items) == []


def test_excel_board_writes_inspection_url_into_empty_d6_only(tmp_path):
    board = ExcelBoard(book_with_domains(tmp_path / "작업.xlsx"), tmp_path / "out.xlsx")
    items = {t.sheet: t for t in board.targets()}
    empty, filled = items["02_빈시트"], items["03_다른사이트"]
    assert empty.d6.startswith("예)") is True        # placeholder is not empty: keep it
    empty.save_domain = filled.save_domain = "https://new.example.com/"
    board.fill(empty, [["가"]], False, tmp_path)
    board.fill(filled, [["나"]], False, tmp_path)
    wb = load_workbook(board.out)
    assert wb["02_빈시트"]["D6"].value.startswith("예)")
    assert wb["03_다른사이트"]["D6"].value == "https://other.example.org"
    wb2 = load_workbook(tmp_path / "작업.xlsx")
    wb2["02_빈시트"]["D6"] = None
    wb2.save(tmp_path / "작업2.xlsx")
    board2 = ExcelBoard(tmp_path / "작업2.xlsx", tmp_path / "out2.xlsx")
    t = {x.sheet: x for x in board2.targets()}["02_빈시트"]
    t.save_domain = "https://new.example.com/"
    board2.fill(t, [["다"]], False, tmp_path)
    assert load_workbook(board2.out)["02_빈시트"]["D6"].value == "https://new.example.com/"


def test_console_text_has_no_ambiguous_symbols():
    from menu_crawler import ui
    text = ui.plain("메뉴 3개 · 방문 2개 → ‘29’ ※ ★ └ ■ ▶ …")
    import unicodedata
    assert not [c for c in text if ord(c) > 127 and unicodedata.east_asian_width(c) == "A"]
    assert ui.width(ui.fit("가나다라마바사", 9)) == 9


def test_no_page_or_time_limit_by_default():
    from menu_crawler import cli
    from menu_crawler.models import CrawlConfig
    config = CrawlConfig(url="https://example.com")
    assert (config.max_pages, config.max_seconds) == (0, 0)
    args = cli.parse_args([])
    assert (args.max_pages, args.max_seconds) == (0, 0)
    assert cli.limits(0, 0) == "페이지 수·시간 제한 없음"


def test_d6_notes_are_not_addresses():
    assert diag.domains("메인 https://www.a.go.kr 입니다.") == ["https://www.a.go.kr"]
    assert diag.domains("www.a.com(운영) 외 2개.") == ["https://www.a.com"]
    assert diag.domains("v1.0 admin@a.com www.b.com") == ["https://www.b.com"]
    assert diag.address("v1.8") == "" and diag.address("n") == "" and diag.address("01_www.a.com") == ""
    t = diag.Target("01_a", "a", d6="정부.한국")
    assert diag.match_url("https://xn--or3bn7q.xn--3e0b707e/main", [t]) == [t]


def test_excel_fills_after_last_used_column_not_into_gaps(tmp_path):
    path = book_with_domains(tmp_path / "d.xlsx")
    wb = load_workbook(path)
    ws = wb["02_빈시트"]
    ws["H13"], ws["J13"] = "가", "다"            # I left empty on purpose
    wb.save(path)
    p = diag.plan(path, "02_빈시트", [["새1"], ["새2"]])
    assert [diag.L(c) for c, _ in p.to_write] == ["K", "L"]


def test_paste_board_same_sheet_twice_continues_after_first_block(tmp_path):
    from menu_crawler.boards import PasteBoard
    board = PasteBoard(str(book_with_domains(tmp_path / "d.xlsx")))
    target = {t.sheet: t for t in board.targets()}["01_테스트"]
    a = board.fill(target, [["가"], ["나"]], False, tmp_path)
    b = board.fill(target, [["나"], ["다"]], False, tmp_path)
    assert (a.where, b.where, b.written, b.duplicates) == ("J13:K15", "L13:L15", 1, 1)
    assert len(list(tmp_path.glob("구글시트붙여넣기_*[0-9].tsv"))) == 2


def test_calcpr_goes_after_external_references(tmp_path):
    import zipfile
    from lxml import etree
    src = book_with_domains(tmp_path / "d.xlsx")
    bad = tmp_path / "noc.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(bad, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                root = etree.fromstring(data)
                for e in list(root):
                    if e.tag.endswith(("calcPr", "definedNames")):
                        root.remove(e)
                sheets = [e for e in root if e.tag.endswith("sheets")][0]
                sheets.addnext(etree.Element(sheets.tag.replace("sheets", "externalReferences")))
                data = etree.tostring(root)
            zout.writestr(item, data)
    out = diag.write(bad, tmp_path / "o.xlsx", diag.plan(bad, "02_빈시트", [["가"]]))
    with zipfile.ZipFile(out) as z:
        tags = [e.tag.split("}")[1] for e in etree.fromstring(z.read("xl/workbook.xml"))]
    assert tags.index("calcPr") == tags.index("externalReferences") + 1


def test_token_links_are_one_page():
    from menu_crawler.urls import page_shape
    assert page_shape("https://a.com/m.do?menuNo=1&_csrf=ab12cd34") == page_shape("https://a.com/m.do?menuNo=1&_csrf=99ff00ee")
    assert page_shape("https://a.com/m.do?menuNo=1") != page_shape("https://a.com/m.do?menuNo=2")
    assert page_shape("https://a.com/x?_=1731400000000") == page_shape("https://a.com/x?_=1731400099999")


def test_closed_input_stops_instead_of_looping(monkeypatch):
    from menu_crawler import cli

    def eof(prompt=""):
        raise EOFError
    monkeypatch.setattr("builtins.input", eof)
    with pytest.raises(SystemExit):
        cli.main([])


def test_tidy_removes_parents_partial_copies_and_controls_and_puts_login_first():
    menus = [["Corporate"], ["Corporate", "Cash Management"], ["Customer Center"],
             ["Customer Center", "User Guide", "Service Hours"], ["User Guide", "Service Hours"], ["User Guide"],
             ["Customer Center", "FAQ"], ["FAQ"], ["About Us"], ["Open Whole Menu"],
             ["Login"], ["Login", "Forgot User ID", "User ID issued at the branch"], ["Login", "Forgot Password"]]
    assert diag.tidy(menus) == [
        ["Login", "Forgot User ID", "User ID issued at the branch"], ["Login", "Forgot Password", ""],
        ["Corporate", "Cash Management", ""], ["Customer Center", "User Guide", "Service Hours"],
        ["Customer Center", "FAQ", ""], ["About Us", "", ""]]


def test_mapping_levels_for_dev_network_sub_urls():
    root = diag.Target("01_root", d6="dev.example.com")
    app = diag.Target("02_app", d6="https://shared.example.com/global.app")
    card = diag.Target("03_card", d6="https://shared.example.com/card/index.jsp")
    parent = diag.Target("04_parent", d6="example.org")
    items = [root, app, card, parent]
    assert diag.match_levels("https://dev.example.com/global.app/login.do", items) == (4, [root])
    assert diag.match_levels("https://shared.example.com/global.app/main.do", items) == (4, [app])   # context path kept
    assert diag.match_levels("https://shared.example.com/card/list.do", items) == (4, [card])
    assert diag.match_levels("https://shared.example.com/", items) == (3, [app, card])            # above both: ask
    assert diag.match_levels("https://shared.example.com/other", items)[0] == 2
    assert diag.match_levels("https://u2lusdev.example.org/x", items) == (1, [parent])
    assert diag.match_url("https://u2lusdev.example.org/x", items) == []                          # weak: ask first


def test_crawl_scope_only_when_a_host_is_shared():
    a = diag.Target("01_a", d6="https://shared.example.com/bank")
    b = diag.Target("02_b", d6="https://shared.example.com/card/index.jsp")
    c = diag.Target("03_c", d6="https://alone.example.com/app")
    items = [a, b, c]
    assert diag.crawl_scope(a, "https://shared.example.com/bank/main.do", items) == "/bank"
    assert diag.crawl_scope(b, "https://shared.example.com/card/", items) == "/card"
    assert diag.crawl_scope(c, "https://alone.example.com/app/x", items) == ""     # alone on its host: whole site


def test_sheet_address_forms_and_last_sheet(tmp_path, monkeypatch):
    from menu_crawler import state
    ref = gsheet.sheet_ref("https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789/edit?gid=1003960627#gid=1003960627")
    assert ref == "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789/edit"
    assert gsheet.sheet_ref("1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789") == ref
    assert gsheet.sheet_ref("https://example.com/x") == "" and gsheet.sheet_ref("뉴") == ""
    monkeypatch.delenv("MENU_CRAWLER_NO_STATE")
    monkeypatch.setattr(state, "FILE", tmp_path / "last_run.json")
    state.remember_sheet(ref, "진단표")
    assert state.last_sheet() == (ref, "진단표")


def test_connections_saved_per_account_and_sheet_bound_to_one(tmp_path, monkeypatch):
    from menu_crawler import cli, state
    monkeypatch.setattr(gsheet, "CONFIG", tmp_path / "google_sheet.json")
    monkeypatch.delenv("MENU_CRAWLER_NO_STATE")
    monkeypatch.setattr(state, "FILE", tmp_path / "last_run.json")
    (tmp_path / "google_sheet.json").write_text('{"url": "https://script.google.com/a", "token": "old"}')   # old format
    assert [c["name"] for c in gsheet.connections()] == ["연결 1"]
    gsheet.save_config("https://script.google.com/b", "k2", "회사 계정")
    gsheet.save_config("https://script.google.com/a", "k1", "개인 계정")                                  # replaces same URL
    assert [(c["name"], c["token"]) for c in gsheet.connections()] == [("회사 계정", "k2"), ("개인 계정", "k1")]
    team = "https://docs.google.com/spreadsheets/d/TEAMTEAMTEAMTEAMTEAMTEAMTEAM12/edit"
    answers = iter(["1"])
    monkeypatch.setattr("builtins.input", lambda prompt="": next(answers))
    assert cli.pick_connection(team, True)["name"] == "회사 계정"        # asked once
    assert cli.pick_connection(team, True)["name"] == "회사 계정"        # remembered, not asked again
    assert len(gsheet.remove_connection(1)) == 1
