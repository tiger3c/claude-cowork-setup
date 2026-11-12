from openpyxl import load_workbook

from menu_crawler import cli
from menu_crawler.models import CrawlConfig, Menu, Report

from test_modes import site  # noqa: F401 - pytest fixture
from test_foundation import launch_options


def report(menus):
    return Report(config=CrawlConfig(url="https://example.test").model_dump(), menus=menus, state="complete").to_dict()


def test_google_sheet_block_layout_and_formula_safety():
    menus = [Menu("a", ["조회", "계좌조회", "거래내역"], "", "", "메뉴 DOM", status="수집됨"),
             Menu("b", ["=공지\t사항"], "", "", "메뉴 DOM", status="수집됨")]
    rows = cli.paste_blocks(report(menus))[0].split("\r\n")
    assert rows[:3] == ["조회\t'=공지 사항", "계좌조회\t", "거래내역\t"]


def test_google_sheet_blocks_split_at_250():
    menus = [Menu(str(i), [f"메뉴 {i}"], "", "", "메뉴 DOM", status="수집됨") for i in range(251)]
    blocks = cli.paste_blocks(report(menus))
    assert len(blocks) == 2
    assert blocks[0].split("\r\n")[0].count("\t") == 249
    assert blocks[1].startswith("메뉴 250")


def test_non_interactive_public_run(site, tmp_path, monkeypatch):  # noqa: F811
    if launch_options():
        monkeypatch.setenv("MENU_CRAWLER_BROWSER_PATH", launch_options()["executable_path"])
    out = tmp_path / "result.xlsx"
    cli.main([site, "--mode", "public", "--headless", "--browser", "chromium", "-y",
              "--wait-ms", "0", "--delay-ms", "0", "--out", str(out)])
    wb = load_workbook(out)
    names = [c.value for c in wb["진단메뉴_01"][13][7:]]
    assert "고객센터" in names and "계좌" not in names
    tsv = (tmp_path / "result_구글시트붙여넣기_01.tsv").read_bytes().decode("utf-8")
    assert tsv.split("\r\n")[1].split("\t")[names.index("고객센터")] == "공지사항"
