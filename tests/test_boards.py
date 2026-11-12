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
