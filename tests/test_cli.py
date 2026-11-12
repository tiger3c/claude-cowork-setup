import os

import pytest
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
    tsv = next(tmp_path.glob("구글시트붙여넣기_*[0-9].tsv")).read_bytes().decode("utf-8")
    assert tsv.split("\r\n")[1].split("\t")[names.index("고객센터")] == "공지사항"


def test_save_never_overwrites(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    data = report([Menu("a", ["메뉴"], "", "", "메뉴 DOM", status="수집됨")])
    data["config"]["site_name"] = 'A/B:"사이트". '
    first, second = cli.save(data, ""), cli.save(data, "")
    assert first != second and first.exists() and second.exists()
    assert first.parent.name == "결과" and "/" not in first.name and ":" not in first.name


class BrokenJob:
    """Engine dies (e.g. the browser driver was killed) after collecting some menus."""

    def __init__(self, config):
        import threading
        self.stop = threading.Event()
        self.thread = threading.Thread(target=lambda: None)
        self.thread.start()
        menus = [Menu("a", ["조회"], "https://example.test/a", "", "메뉴 DOM", status="수집됨")]
        self.data = {**report(menus), "state": "error", "message": "수집 브라우저가 닫혔습니다."}

    def snapshot(self):
        return self.data

    def command(self, *args, **kw):
        return 1

    def close(self):
        pass


def test_engine_error_still_saves_partial_result(tmp_path, monkeypatch):
    monkeypatch.setattr(cli, "Job", BrokenJob)
    out = tmp_path / "partial.xlsx"
    cli.main(["https://example.test", "--mode", "public", "-y", "--out", str(out)])
    assert load_workbook(out)["진단메뉴_01"]["H13"].value == "조회"


@pytest.mark.skipif(os.name == "nt", reason="clipboard is available on Windows")
def test_clipboard_reports_unavailable_off_windows():
    assert cli.copy_to_clipboard("a\tb\r\n") is False
