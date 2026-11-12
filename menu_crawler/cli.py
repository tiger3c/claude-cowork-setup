"""Interactive command-line runner.

Flow: site URL -> collection mode -> only the questions that mode needs -> crawl -> Excel.
Every question can also be answered with a command-line option for repeat runs.
"""
import argparse
import getpass
import os
import subprocess
import sys
import tempfile
import time
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit

from pydantic import ValidationError

from .crawler import Job
from .export import make_workbook, select_menus
from .models import MODES, CrawlConfig

BUSY = {"starting", "crawling", "capturing"}
MODE_KEYS = {"1": "public", "2": "login", "3": "compare"}
MODE_HELP = {
    "public": "로그인하지 않은 방문자에게 보이는 공개 메뉴만 수집",
    "login": "브라우저 창에서 직접 로그인한 뒤 로그인 상태의 메뉴를 수집",
    "compare": "공개 메뉴 수집 → 로그인 → 다시 수집해 '로그인 후에만 보이는 메뉴'를 구분",
}


# ---------- 화면 출력 ----------

def line(char="─"):
    print(char * 64)


def title(text):
    print()
    line("═")
    print(f"  {text}")
    line("═")


def ask(prompt, default=""):
    suffix = f" [{default}]" if default else ""
    try:
        value = input(f"{prompt}{suffix}: ").strip()
    except EOFError:
        value = ""
    return value or default


def ask_yes(prompt, default=False):
    mark = "Y/n" if default else "y/N"
    value = ask(f"{prompt} ({mark})").lower()
    return default if not value else value in {"y", "yes", "예", "ㅇ"}


def ask_int(prompt, default, low, high):
    while True:
        value = ask(prompt, str(default))
        if value.isdigit() and low <= int(value) <= high:
            return int(value)
        print(f"  {low}~{high} 사이 숫자를 입력하세요.")


# ---------- 입력 받기 ----------

def ask_url(initial):
    url = initial
    while True:
        if not url:
            url = ask("수집할 사이트 주소 (예: https://www.example.com)")
        if url and "://" not in url:
            url = "https://" + url
        try:
            return CrawlConfig(url=url).url
        except ValidationError:
            print("  http:// 또는 https://로 시작하는 주소를 입력하세요.")
            url = ""


def ask_mode(initial):
    if initial:
        return initial
    print()
    print("수집 방식을 선택하세요.")
    for key, mode in MODE_KEYS.items():
        print(f"  {key}) {MODES[mode]:<16} - {MODE_HELP[mode]}")
    while True:
        value = ask("번호", "1")
        if value in MODE_KEYS:
            return MODE_KEYS[value]
        print("  1, 2, 3 중에서 선택하세요.")


def build_config(args):
    title("웹 메뉴 수집기")
    url = ask_url(args.url)
    mode = ask_mode(args.mode)
    interactive = not args.yes
    host = urlsplit(url).hostname or "메뉴 수집"
    name = args.name or (ask("대상 이름 (엑셀 요약·파일명에 사용)", host) if interactive else host)

    login_url = args.login_url or ""
    if mode != "public" and interactive and not login_url:
        login_url = ask("로그인 페이지 주소 (없으면 Enter: 사이트 주소에서 직접 이동)")

    settings = {"max_pages": args.max_pages, "max_seconds": args.max_seconds, "max_hops": args.max_hops,
                "wait_ms": args.wait_ms, "delay_ms": args.delay_ms}
    if interactive and ask_yes("고급 설정(방문 페이지 수·시간·간격)을 바꿀까요?"):
        settings["max_pages"] = ask_int("  최대 방문 페이지 수", settings["max_pages"], 1, 2000)
        settings["max_seconds"] = ask_int("  최대 수집 시간(초)", settings["max_seconds"], 10, 7200)
        settings["max_hops"] = ask_int("  최대 링크 이동 횟수", settings["max_hops"], 1, 20)
        settings["delay_ms"] = ask_int("  페이지 간 간격(ms, 서버 부담을 줄이려면 늘리기)", settings["delay_ms"], 0, 5000)
        hosts = ask("  추가 허용 호스트 (쉼표로 구분, 없으면 Enter)")
        args.extra_host = args.extra_host + [h.strip() for h in hosts.split(",") if h.strip()]
        args.selector = ask("  메뉴 영역 CSS 선택자 (없으면 Enter)", args.selector)

    headless = args.headless and mode == "public"
    if args.headless and mode != "public":
        print("  ※ 로그인이 필요한 방식이라 브라우저 창을 띄워서 실행합니다.")

    try:
        config = CrawlConfig(url=url, mode=mode, site_name=name[:120], login_url=login_url, extra_hosts=args.extra_host,
                             browser=args.browser, menu_selector=args.selector, headless=headless, **settings)
    except ValidationError as exc:
        sys.exit("입력값 오류: " + "; ".join(e["msg"] for e in exc.errors()))

    print()
    line()
    print(f"  대상       : {config.site_name}")
    print(f"  사이트 주소: {config.url}")
    print(f"  수집 방식  : {MODES[config.mode]}")
    if config.mode != "public":
        print(f"  로그인 주소: {config.login_url or '(사이트 주소에서 직접 이동)'}")
    print(f"  제한       : 페이지 {config.max_pages}개 · {config.max_seconds}초 · 이동 {config.max_hops}회")
    line()
    if not args.yes and not ask_yes("수집 대상 사이트에 대한 점검 권한(계약·승인)을 확인했고, 이대로 시작할까요?", True):
        sys.exit("취소했습니다.")
    return config


# ---------- 진행 ----------

def wait(job, quiet=False):
    """Wait while the engine is busy, printing one progress line. Ctrl+C asks the engine to stop."""
    last = ""
    try:
        while True:
            data = job.snapshot()
            if data["state"] not in BUSY:
                break
            if not quiet:
                pages = f"{data['visited_count']}/{data['config']['max_pages']}"
                current = data.get("current_url", "")
                text = f"  [{pages}] 메뉴 {len(data['menus'])}개 · {current}"[:110]
                if text != last:
                    print("\r" + text.ljust(110), end="", flush=True)
                    last = text
            time.sleep(0.4)
    except KeyboardInterrupt:
        print("\n  중지 요청: 현재 페이지까지만 확인하고 멈춥니다...")
        job.stop.set()
        while job.snapshot()["state"] in BUSY:
            time.sleep(0.3)
    if last:
        print()
    data = job.snapshot()
    if data["state"] == "error":
        sys.exit("실행 오류: " + data["message"])
    return data


def run_command(job, name, **args):
    """Queue a quick command (goto/fill) and wait until the engine has finished it."""
    return job.wait_done(job.command(name, **args))


def crawl(job, phase):
    job.command("crawl", phase=phase)
    data = wait(job)
    print("  " + data["message"])
    if data.get("login_check"):
        print("  ⚠ " + data["login_check"])
    return data


def login_step(job, config, username=None):
    title("로그인 단계")
    print("  1. 열린 브라우저 창에서 로그인하세요. 보안문자·OTP·추가 인증까지 마칩니다.")
    print("  2. 로그인된 화면이 보이면 여기서 Enter를 누르세요.")
    print("  ※ 아이디·비밀번호는 파일·엑셀·기록에 저장하지 않습니다.")
    print("  " + run_command(job, "goto", url=config.login_url or config.url)["message"])
    if username is not None:
        autofill(job, username)
    while True:
        print()
        print("  [Enter] 로그인 완료 → 수집 시작")
        print("  [a] 아이디·비밀번호 자동 입력 (로그인 버튼은 직접 누름)")
        print("  [o] 로그인 화면 다시 열기")
        print("  [c] 지금 열린 화면의 메뉴만 추가")
        print("  [q] 로그인 없이 종료하고 지금까지 결과 저장")
        choice = ask("선택").lower()
        if choice == "":
            return True
        if choice == "a":
            autofill(job, ask("  아이디"))
        elif choice == "o":
            print("  " + run_command(job, "goto", url=config.login_url or config.url)["message"])
        elif choice == "c":
            job.command("capture", phase="after")
            print("  " + wait(job, quiet=True)["message"])
        elif choice == "q":
            return False


def autofill(job, username):
    password = getpass.getpass("  비밀번호 (입력 내용은 화면에 표시되지 않음): ")
    if not password:
        print("  비밀번호가 비어 있어 건너뜁니다.")
        return
    ticket = job.command("fill", username=username, password=password)
    del password
    print("  " + job.wait_done(ticket)["message"])


def summary(data):
    title("수집 결과")
    menus = data["menus"]
    ready = sum(m["ready"] for m in menus)
    print(f"  발견 메뉴 {len(menus)}개 · 진단 시트 붙여넣기 대상 {ready}개 · 방문 페이지 {data['visited_count']}개")
    for status, count in Counter(m["status"] for m in menus).most_common():
        print(f"    - {status}: {count}")
    vis = Counter(m["visibility"] for m in menus if m.get("visibility"))
    if vis:
        print(f"  로그인 전 공개 {vis['로그인 전 공개']}개 · 로그인 후 전용 {vis['로그인 후 전용']}개")
    open_issues = [x for x in data["issues"] if not x.get("resolved")]
    if open_issues:
        print(f"  확인 필요 {len(open_issues)}건 (엑셀 '미수집_확인필요' 시트 참고)")
        for item in open_issues[:5]:
            print(f"    - [{item['phase']}] {item['kind']}: {item['url']}"[:110])
        if len(open_issues) > 5:
            print(f"    ... 외 {len(open_issues) - 5}건")
    if data["limit_reached"]:
        print("  ⚠ 수집 제한(페이지 수·시간·중지)에 도달했습니다. 필요하면 고급 설정을 늘려 다시 실행하세요.")
    tree(menus)


def tree(menus, limit=40):
    shown = [m for m in menus if m["ready"]][:limit]
    if not shown:
        return
    print()
    print("  메뉴 구조 미리보기")
    prev = []
    for m in shown:
        for depth, name in enumerate(m["path"][:3]):
            if prev[:depth + 1] != m["path"][:depth + 1]:
                mark = " ★" if depth == len(m["path"]) - 1 and m.get("visibility") == "로그인 후 전용" else ""
                print("    " + "  " * depth + ("└ " if depth else "■ ") + name + mark)
        prev = m["path"][:3]
    rest = sum(m["ready"] for m in menus) - len(shown)
    if rest > 0:
        print(f"    ... 외 {rest}개는 엑셀에서 확인하세요.")
    if any(m.get("visibility") == "로그인 후 전용" for m in shown):
        print("    (★ 로그인 후에만 보이는 메뉴)")


def save(data, out):
    path = Path(out) if out else Path("결과") / (
        "".join("_" if c in '\\/:*?"<>|' else c for c in data["config"]["site_name"])
        + f"_메뉴목록_{datetime.now():%Y%m%d_%H%M}.xlsx")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(make_workbook(data))
    return path.resolve()


# ---------- 구글 시트 붙여넣기 ----------

def sheet_text(value):
    """One cell for a Google Sheets paste: no tabs/newlines, never parsed as a formula."""
    value = " ".join(str(value).split())
    return "'" + value if value[:1] in {"=", "+", "-", "@"} else value


def paste_blocks(data, size=250):
    """Rows = 1/2/3 Depth, columns = menus; the same layout as the H13:..15 range."""
    menus = select_menus(data)
    blocks = []
    for start in range(0, len(menus), size):
        chunk = menus[start:start + size]
        rows = [[sheet_text(m["depths"][d]) for m in chunk] for d in range(3)]
        blocks.append("\r\n".join("\t".join(r) for r in rows) + "\r\n")
    return blocks


def copy_to_clipboard(text):
    if os.name != "nt":
        return False
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".txt", delete=False) as f:
        f.write(text)
        temp = f.name
    try:
        command = f"Get-Content -Raw -Encoding UTF8 -LiteralPath '{temp}' | Set-Clipboard"
        return subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
                              capture_output=True, timeout=20).returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False
    finally:
        os.unlink(temp)


def google_sheet_step(data, excel_path, interactive):
    blocks = paste_blocks(data)
    if not blocks:
        print("  붙여넣을 메뉴가 없습니다. 엑셀의 '미수집_확인필요' 시트를 확인하세요.")
        return
    files = []
    for i, block in enumerate(blocks, 1):
        path = excel_path.with_name(f"{excel_path.stem}_구글시트붙여넣기_{i:02}.tsv")
        path.write_text(block, encoding="utf-8")
        files.append(path)
    title("구글 시트에 붙여넣기")
    print(f"  진단 메뉴 {len(select_menus(data))}개 · {len(blocks)}묶음 (묶음당 최대 250개 = H~IW열)")
    print("  방법: 구글 시트 진단표에서 H13 셀 클릭 → Ctrl+V")
    print("        (1행=1 Depth, 2행=2 Depth, 3행=3 Depth로 H13:15에 가로로 들어갑니다)")
    for i, path in enumerate(files, 1):
        print(f"  붙여넣기 파일 {i}: {path}")
    if not interactive:
        return
    if os.name != "nt":
        print("  ※ 이 PC에서는 자동 복사를 지원하지 않습니다. 위 .tsv 파일 내용을 복사해 붙여넣으세요.")
        return
    for i, block in enumerate(blocks, 1):
        label = f"{i}/{len(blocks)}묶음" + (" (다음 진단 시트용)" if i > 1 else "")
        if not ask_yes(f"  {label}을 클립보드에 복사할까요?", True):
            continue
        if copy_to_clipboard(block):
            print("  ✔ 복사했습니다. 구글 시트 H13을 클릭하고 Ctrl+V 하세요.")
        else:
            print("  복사에 실패했습니다. 위 .tsv 파일을 메모장으로 열어 전체 복사하세요.")
        if i < len(blocks):
            ask("  붙여넣었으면 Enter")


# ---------- 실행 ----------

def parse_args(argv=None):
    p = argparse.ArgumentParser(prog="메뉴수집", description="웹사이트 메뉴를 1·2·3 Depth로 수집해 엑셀로 정리합니다. 옵션 없이 실행하면 하나씩 물어봅니다.")
    p.add_argument("url", nargs="?", help="수집할 사이트 주소")
    p.add_argument("--mode", choices=["public", "login", "compare"], help="public=로그인 전, login=로그인 후, compare=전·후 비교")
    p.add_argument("--login-url", default="", help="로그인 페이지 주소")
    p.add_argument("--name", default="", help="대상 이름")
    p.add_argument("--out", default="", help="저장할 엑셀 파일 경로 (기본: 결과 폴더)")
    p.add_argument("--username", help="로그인 단계에서 자동 입력할 아이디 (비밀번호는 실행 중에 따로 물어봄)")
    p.add_argument("--max-pages", type=int, default=300)
    p.add_argument("--max-seconds", type=int, default=1200)
    p.add_argument("--max-hops", type=int, default=6)
    p.add_argument("--wait-ms", type=int, default=700)
    p.add_argument("--delay-ms", type=int, default=300)
    p.add_argument("--extra-host", action="append", default=[], help="추가 허용 호스트 (여러 번 지정 가능)")
    p.add_argument("--selector", default="", help="메뉴 영역 CSS 선택자")
    p.add_argument("--browser", choices=["auto", "msedge", "chrome", "chromium"], default="auto")
    p.add_argument("--headless", action="store_true", help="브라우저 창 없이 실행 (로그인 전 수집만)")
    p.add_argument("-y", "--yes", action="store_true", help="확인 질문 없이 바로 실행 (점검 권한 확인 완료로 간주)")
    p.add_argument("--no-open", action="store_true", help="저장 후 엑셀을 자동으로 열지 않음")
    return p.parse_args(argv)


def finish(data, args):
    summary(data)
    path = save(data, args.out)
    print()
    print(f"  엑셀 저장: {path}")
    google_sheet_step(data, path, not args.yes)
    if os.name == "nt" and not args.no_open and not args.yes and ask_yes("전체 메뉴 엑셀 파일을 열까요?", False):
        os.startfile(path)  # noqa: S606 - Windows only, local result file


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    args = parse_args(argv)
    config = build_config(args)

    title(MODES[config.mode])
    print("  브라우저를 여는 중입니다...")
    job = Job(config)
    data = None
    try:
        data = wait(job, quiet=True)
        print("  " + data["message"])
        if config.mode in {"public", "compare"}:
            if config.mode == "compare":
                print("\n  [1단계] 로그인 전 공개 메뉴 수집")
            data = crawl(job, "before")
        if config.mode in {"login", "compare"} and not job.stop.is_set():
            if login_step(job, config, args.username):
                print("\n  [로그인 후] 메뉴 수집")
                data = crawl(job, "after")
            else:
                data = job.snapshot()
        finish(data, args)
    except KeyboardInterrupt:
        print("\n  중단했습니다.")
        job.stop.set()
        data = job.snapshot()
        if data["menus"]:
            print("  지금까지 수집한 결과를 저장합니다.")
            args.yes = True  # no more questions after Ctrl+C
            finish(data, args)
    finally:
        job.close()
        job.thread.join(timeout=10)


if __name__ == "__main__":
    main()
