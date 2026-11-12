"""Writing crawled menus into a diagnosis workbook (진단 시트 13-15행)."""
import zipfile

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference

from menu_crawler import cli
from menu_crawler import workbook as diag
from menu_crawler.export import make_workbook
from menu_crawler.models import CrawlConfig, Menu, Report


def diagnosis_book(path):
    """A small stand-in for the real template: summary sheet with a chart, sample and two site sheets."""
    wb = Workbook()
    summary = wb.active
    summary.title = "사이트목록"
    for r, (k, v) in enumerate([("01_테스트", 3), ("02_빈시트", 0)], 5):
        summary.cell(r, 7, k)
        summary.cell(r, 4, v)
    chart = BarChart()
    chart.add_data(Reference(summary, min_col=4, min_row=5, max_row=6))
    summary.add_chart(chart, "J2")
    for name in ("00_Sample", "01_테스트", "02_빈시트"):
        ws = wb.create_sheet(name)
        ws["A13"], ws["A14"], ws["A15"] = "1 Depth", "2 Depth", "3 Depth"
        ws["H16"] = '=MID(IF(LEN(TRIM(H13))," > "&TRIM(H13),""),4,32767)'
        ws["D72"] = '=COUNTIF(H13:IW13,"<>")'
    t = wb["01_테스트"]
    t["H13"], t["H14"], t["H20"] = "기존", "하위", "양호"    # existing menu with a result
    t["I20"] = "취약"                                         # result without a menu: column is occupied
    wb["00_Sample"]["H13"] = "예시"
    wb.save(path)
    return path


def test_plan_uses_only_empty_columns_and_skips_existing(tmp_path):
    book = diagnosis_book(tmp_path / "diag.xlsx")
    p = diag.plan(book, "01_테스트", [["기존", "하위", ""], ["조회", "계좌조회", "거래내역"], ["이체"], ["이체"]])
    assert p.existing == 1
    assert [c for c, _ in p.to_write] == [diag.CI("J"), diag.CI("K")]      # H and I are occupied
    assert p.duplicates == [["기존", "하위", ""], ["이체", "", ""]]
    assert diag.Workbook(book).site_sheets() == ["01_테스트", "02_빈시트"]   # 00_Sample is never offered


def test_write_changes_only_menu_cells_and_keeps_everything_else(tmp_path):
    book = diagnosis_book(tmp_path / "diag.xlsx")
    out = tmp_path / "diag_out.xlsx"
    p = diag.plan(book, "01_테스트", [["조회", "계좌조회", "거래내역"], ["=SUM(1)", "", ""]])
    diag.write(book, out, p)
    wb = load_workbook(out)
    t = wb["01_테스트"]
    assert (t["J13"].value, t["J14"].value, t["J15"].value) == ("조회", "계좌조회", "거래내역")
    assert t["K13"].value == "=SUM(1)" and t["K13"].data_type == "s"         # stored as text, never a formula
    assert (t["H13"].value, t["H20"].value, t["I20"].value) == ("기존", "양호", "취약")
    assert t["H16"].value.startswith("=MID(")
    with zipfile.ZipFile(book) as a, zipfile.ZipFile(out) as b:
        changed = [n for n in a.namelist() if a.read(n) != b.read(n)]
        assert sorted(changed) == ["xl/workbook.xml", "xl/worksheets/sheet3.xml"]   # target sheet + recalc flag
        assert any("chart" in n for n in b.namelist())                              # chart kept
        assert b'fullCalcOnLoad="1"' in b.read("xl/workbook.xml")


def test_overflow_and_no_overwrite(tmp_path):
    book = diagnosis_book(tmp_path / "diag.xlsx")
    menus = [[f"메뉴 {i}"] for i in range(260)]
    p = diag.plan(book, "01_테스트", menus)
    assert len(p.to_write) == 248 and len(p.overflow) == 12                # 250 columns - 2 occupied
    with pytest.raises(ValueError):
        diag.write(book, book, p)


def test_cli_from_result_into_diagnosis_sheet(tmp_path):
    menus = [Menu(str(i), path, f"https://example.test/{i}", "", "메뉴 DOM", status="수집됨")
             for i, path in enumerate([["조회", "계좌조회"], ["이체"], ["고객센터", "공지사항", "목록"]])]
    report = Report(config=CrawlConfig(url="https://example.test").model_dump(), menus=menus, state="complete").to_dict()
    result = tmp_path / "result.xlsx"
    result.write_bytes(make_workbook(report))
    assert diag.menus_from_result(result) == [["조회", "계좌조회", ""], ["이체", "", ""], ["고객센터", "공지사항", "목록"]]
    book = diagnosis_book(tmp_path / "diag.xlsx")
    out = tmp_path / "filled.xlsx"
    cli.main(["--from-result", str(result), "--workbook", str(book), "--sheet", "02_빈시트", "--workbook-out", str(out), "-y"])
    ws = load_workbook(out)["02_빈시트"]
    assert [ws.cell(13, c).value for c in range(8, 11)] == ["조회", "이체", "고객센터"]
    assert ws["J15"].value == "목록"


def test_control_characters_xlsm_and_wrong_result_file(tmp_path):
    book = diagnosis_book(tmp_path / "diag.xlsx")
    p = diag.plan(book, "01_테스트", [["메뉴\x0b이름\x1f", "하위\x00메뉴"]])
    out = diag.write(book, tmp_path / "out.xlsx", p)
    ws = load_workbook(out)["01_테스트"]
    assert (ws["J13"].value, ws["J14"].value) == ("메뉴 이름", "하위 메뉴")
    xlsm = tmp_path / "diag.xlsm"
    xlsm.write_bytes(book.read_bytes())
    assert diag.write(xlsm, tmp_path / "out.xlsm", diag.plan(xlsm, "02_빈시트", [["조회"]])).suffix == ".xlsm"
    with pytest.raises(ValueError):
        diag.menus_from_result(book)             # a diagnosis workbook is not a crawl result


def test_unverified_navigation_menus_can_be_included(tmp_path):
    menus = [Menu("a", ["조회"], "https://example.test/a", "", "메뉴 DOM", status="수집됨"),
             Menu("b", ["마이페이지"], "https://example.test/my", "", "메뉴 DOM", status="로그인 필요"),
             Menu("c", ["로그아웃"], "https://example.test/logout", "", "메뉴 DOM", status="동작 확인 필요"),
             Menu("d", ["이벤트"], "https://example.test/e", "", "일반 링크", status="로그인 필요", hierarchy_review=True)]
    report = Report(config=CrawlConfig(url="https://example.test").model_dump(), menus=menus, state="complete").to_dict()
    confirmed, unverified = diag.menus_from_report(report, include_unverified=True)
    assert confirmed == [["조회", "", ""]]
    assert unverified == [(["마이페이지", "", ""], "로그인 필요")]     # logout and body links stay out
    result = tmp_path / "result.xlsx"
    result.write_bytes(make_workbook(report))
    assert diag.unverified_from_result(result) == [(["마이페이지", "", ""], "로그인 필요")]
    book = diagnosis_book(tmp_path / "diag.xlsx")
    out = tmp_path / "filled.xlsx"
    cli.main(["--from-result", str(result), "--workbook", str(book), "--sheet", "02_빈시트",
              "--workbook-out", str(out), "--include-unverified", "-y"])
    ws = load_workbook(out)["02_빈시트"]
    assert [ws["H13"].value, ws["I13"].value, ws["J13"].value] == ["조회", "마이페이지", None]
