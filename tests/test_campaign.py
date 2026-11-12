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
    assert cli.split_tokens("1 2, 29 - 34, 뉴욕 지점") == ["1", "2", "29-34", "뉴욕 지점"]
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
