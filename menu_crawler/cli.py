"""Interactive command-line runner.

Flow: site URL -> collection mode -> only the questions that mode needs -> crawl -> Excel.
Every question can also be answered with a command-line option for repeat runs.
"""
import argparse
import getpass
import ctypes
import os
import re
import shutil
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
from . import workbook as diag
from . import gsheet, mybrowser, netcheck, state, ui
from .boards import ExcelBoard, Filled, PasteBoard, SheetsBoard
from .gsheet import SheetError
from .ui import ask, ask_int, ask_yes
from .export import make_workbook, select_menus
from .models import MODES, CrawlConfig
from .urls import normalize_host

BUSY = {"starting", "crawling", "capturing"}
MODE_KEYS = {"1": "auto", "2": "public", "3": "login", "4": "compare"}
MODE_LABELS = {"auto": "자동 (필요할 때만 로그인)", **MODES}
MODE_HELP = {
    "auto": "로그인 없이 먼저 수집하고, 로그인 화면만 보이는 대상은 그때 로그인한 뒤 수집",
    "public": "로그인하지 않은 방문자에게 보이는 공개 메뉴만 수집",
    "login": "브라우저 창에서 직접 로그인한 뒤 로그인 상태의 메뉴를 수집",
    "compare": "공개 메뉴 수집 → 로그인 → 다시 수집해 '로그인 후에만 보이는 메뉴'를 구분",
}


# ---------- 화면 출력 ----------

def line(char="─"):
    print(char * 64)


def title(text):
    ui.header(text)


# ---------- 입력 받기 ----------

def limits(pages, seconds):
    if not pages and not seconds:
        return "페이지 수·시간 제한 없음"
    return f"페이지 {pages or '제한 없음'}{'개' if pages else ''} · 시간 {f'{seconds}초' if seconds else '제한 없음'}"


def ask_url(initial, prompt="수집할 사이트 주소 (예: https://www.example.com)", default=""):
    url = initial
    while True:
        if not url:
            url = ask(prompt, default)
        if url and "://" not in url:
            url = "https://" + url
        try:
            if diag.address(url):
                return CrawlConfig(url=url).url
        except ValidationError:
            pass
        print("  주소 형식이 아닙니다. 예: www.example.com 또는 https://www.example.com")
        url = ""


def ask_mode(initial, default="1"):
    if initial:
        return initial
    last = state.pref("mode")
    default = next((k for k, m in MODE_KEYS.items() if m == last), default)
    print()
    print("  수집 방식을 선택하세요.")
    for key, mode in MODE_KEYS.items():
        print(f"    {key}) {ui.fit(MODE_LABELS[mode], 26)} - {MODE_HELP[mode]}")
    while True:
        value = ask("  번호", default)
        if value in MODE_KEYS:
            state.remember("mode", MODE_KEYS[value])
            return MODE_KEYS[value]
        print("  1~4 중에서 선택하세요.")


def base_settings(args):
    return {"max_pages": args.max_pages, "max_seconds": args.max_seconds, "max_hops": args.max_hops,
            "wait_ms": args.wait_ms, "delay_ms": args.delay_ms, "click_menus": not args.no_click}


def show_settings(args, s):
    rows = [
        ["방문 페이지 수", f"{s['max_pages']}개" if s["max_pages"] else "제한 없음", "대상 하나에서 열어 볼 최대 페이지 수"],
        ["수집 시간", f"{s['max_seconds']}초" if s["max_seconds"] else "제한 없음", "대상 하나에 쓰는 최대 시간"],
        ["링크 깊이", f"{s['max_hops']}단계", "메뉴에서 링크를 몇 번까지 따라 들어갈지"],
        ["페이지 간격", f"{s['delay_ms'] / 1000:g}초", "다음 페이지를 열기 전 쉬는 시간 (서버 부담이 걱정되면 늘림)"],
        ["메뉴 눌러 보기", "켜짐" if s["click_menus"] else "꺼짐", "javascript: 메뉴를 눌러 열리는 화면과 탭까지 확인"],
        ["메뉴 영역", args.selector or "자동으로 찾기", "메뉴를 못 찾을 때만 CSS 선택자 지정 (예: #topMenu)"],
        ["추가 호스트", ", ".join(args.extra_host) or "없음", "메뉴가 다른 주소(예: 별도 로그인 서버)로 이어질 때 함께 수집"],
    ]
    ui.table([("항목", 14), ("지금 값", 14), ("뜻", 0)], rows, indent="    ")


def advanced_settings(args, settings, interactive):
    """Show what each advanced option means and its value; change only when asked. Enter keeps a value."""
    if not interactive:
        return
    print()
    short = (f"페이지 {'제한 없음' if not settings['max_pages'] else settings['max_pages']} · "
             f"시간 {'제한 없음' if not settings['max_seconds'] else str(settings['max_seconds']) + '초'} · "
             f"메뉴 눌러 보기 {'켜짐' if settings['click_menus'] else '꺼짐'} · 메뉴 영역 {args.selector or '자동'}")
    print("  고급 설정: " + ui.dim(short))
    if not ask_yes("  고급 설정을 보거나 바꿀까요? (보통은 그대로)", False):
        return
    show_settings(args, settings)
    print("  " + ui.dim("Enter = 지금 값 그대로"))
    settings["max_pages"] = ask_int("    방문 페이지 수 (0 = 제한 없음)", settings["max_pages"], 0, 100000)
    settings["max_seconds"] = ask_int("    수집 시간, 초 (0 = 제한 없음, 예: 600 = 10분)", settings["max_seconds"], 0, 172800)
    settings["max_hops"] = ask_int("    링크 깊이 (1~20)", settings["max_hops"], 1, 20)
    settings["delay_ms"] = ask_int("    페이지 간격, 밀리초 (300 = 0.3초)", settings["delay_ms"], 0, 5000)
    settings["click_menus"] = ask_yes("    javascript: 메뉴를 눌러 보기", settings["click_menus"])
    args.selector = ask("    메뉴 영역 CSS 선택자 (없으면 Enter = 자동)", args.selector)
    while True:
        hosts = [h.strip() for h in ask("    추가 호스트 (쉼표로 구분, 없으면 Enter, 예: login.example.com)").split(",") if h.strip()]
        try:
            args.extra_host = args.extra_host + [normalize_host(h) for h in hosts]
            break
        except ValueError as exc:
            print(f"  {exc}")
    print("  변경한 설정")
    show_settings(args, settings)


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

    settings = base_settings(args)
    advanced_settings(args, settings, interactive)

    headless = args.headless
    if args.headless and mode != "public":
        print("  ※ 창 없이(--headless) 실행하면 로그인 단계에서 직접 로그인할 수 없습니다.")
    args.auto_login = mode == "auto"

    if mode != "public" and interactive:
        ask_login_window(args)
    if args.my_browser:
        try:
            settings["cdp_url"] = mybrowser.start(url, headless=args.headless)
        except RuntimeError as exc:
            sys.exit(str(exc))
    try:
        config = CrawlConfig(url=url, mode=engine_mode(mode), site_name=name[:120], login_url=login_url, extra_hosts=args.extra_host,
                             browser=args.browser, menu_selector=args.selector, headless=headless, **settings)
    except ValidationError as exc:
        sys.exit("입력값 오류: " + "; ".join(e["msg"] for e in exc.errors()))

    print()
    line()
    print(f"  대상       : {config.site_name}")
    print(f"  사이트 주소: {config.url}")
    print(f"  수집 방식  : {MODE_LABELS['auto' if args.auto_login else config.mode]}")
    if config.mode != "public":
        print(f"  로그인 주소: {config.login_url or '(사이트 주소에서 직접 이동)'}")
    print(f"  제한       : {limits(config.max_pages, config.max_seconds)} · 이동 {config.max_hops}회")
    line()
    if not args.yes and not ask_yes("수집 대상 사이트에 대한 점검 권한(계약·승인)을 확인했고, 이대로 시작할까요?", True):
        sys.exit("취소했습니다.")
    return config


# ---------- 진행 ----------

class EngineError(Exception):
    """The browser engine failed; the partial result is still worth saving."""


def stop_key_pressed():
    """Windows: 'q' stops safely. (Ctrl+C there also reaches the browser driver and can kill it.)"""
    if os.name != "nt":
        return False
    import msvcrt
    pressed = False
    while msvcrt.kbhit():
        pressed |= msvcrt.getwch().lower() in {"q", "ㅂ"}
    return pressed


def wait(job, quiet=False):
    """Wait while the engine is busy, printing one progress line. q or Ctrl+C asks the engine to stop."""
    last = ""
    width = max(40, shutil.get_terminal_size((100, 20)).columns - 1)
    if not quiet and os.name == "nt":
        print("  (중지하려면 q 키를 누르세요. 지금까지의 결과는 저장됩니다.)")
    try:
        while True:
            data = job.snapshot()
            if data["state"] not in BUSY:
                break
            if stop_key_pressed() and not job.stop.is_set():
                print("\n  중지 요청: 현재 페이지까지만 확인하고 멈춥니다...")
                job.stop.set()
            if not quiet:
                cap = data["config"]["max_pages"]
                pages = f"{data['visited_count']}/{cap}" if cap else f"{data['visited_count']}페이지"
                current = data.get("current_url", "")
                text = ui.fit(f"  [{pages}] 메뉴 {len(data['menus'])}개 · {current}", width)
                if text != last:
                    print("\r" + text, end="", flush=True)
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
        raise EngineError(data["message"])
    return data


def run_command(job, name, **args):
    """Queue a quick command (goto/fill) and wait until the engine has finished it."""
    return job.wait_done(job.command(name, **args))


def crawl(job, phase, resume=False):
    job.command("crawl", phase=phase, resume=resume)
    data = wait(job)
    print("  " + data["message"])
    if data.get("login_check"):
        print("  ⚠ " + data["login_check"])
    return data


class Account:
    """Login used for this run only. The password, if the user agrees, is kept in memory until the program ends:
    never in a file, an Excel or a log. It is only ever typed into the same site (registrable domain) it came from."""
    username = ""
    password = ""
    site = ""


def site_of(url):
    host = (urlsplit(url).hostname or "").lower()
    parts = host.split(".")
    if host.replace(".", "").isdigit() or len(parts) <= 2:
        return host
    # example.co.kr / example.com.au: keep three labels when the second-level label is a generic one
    return ".".join(parts[-3:]) if parts[-2] in {"co", "or", "go", "ac", "ne", "re", "pe", "com", "net", "org", "gov", "edu"} else ".".join(parts[-2:])


def login_step(job, config, username=None):
    title("로그인 단계")
    ui.beep()
    ui.window_title("로그인이 필요합니다")
    if config.cdp_url:
        print("  1. 내 브라우저 창(수집기가 연 Edge/Chrome)의 새 탭에서 로그인하세요. 인증서·OTP·보안 프로그램도 그 창에서 처리합니다.")
    else:
        print("  1. 열린 브라우저 창에서 로그인하세요. 보안문자·OTP·추가 인증까지 마칩니다.")
    print("  2. 로그인된 화면이 보이면 여기서 Enter를 누르세요.")
    print("  ※ 아이디·비밀번호는 파일·엑셀·기록에 저장하지 않습니다.")
    print("  " + run_command(job, "goto", url=config.login_url or config.url)["message"])
    if username is not None and not Account.username:
        Account.username = username
    site = site_of(config.login_url or config.url)
    if Account.username and Account.password and Account.site == site:
        print("  " + ui.dim(f"같은 사이트({site})에서 쓴 계정({Account.username})을 입력란에 채웁니다. 로그인 버튼은 직접 누르세요."))
        autofill(job, Account.username, Account.password, site)
    elif username is not None:
        autofill(job, username, site=site)
    while True:
        print()
        print("  [Enter] 로그인 완료 -> 수집 시작")
        print("  [a] 아이디/비밀번호 자동 입력 (로그인 버튼은 직접 누름)")
        print("  [o] 로그인 화면 다시 열기")
        print("  [c] 지금 열린 화면의 메뉴만 추가")
        print("  [p] 로그인 없이 공개 메뉴만 수집 (계정 문제 등으로 로그인이 안 될 때)")
        print("  [q] 로그인 없이 종료하고 지금까지 결과 저장")
        choice = ask("  선택").lower()
        if choice == "":
            return True
        if choice in {"p", "ㅔ"}:
            return "public"
        if choice in {"a", "ㅁ"}:
            Account.username = ask("  아이디", Account.username if Account.site == site else "")
            autofill(job, Account.username, site=site)
        elif choice in {"o", "ㅐ"}:
            print("  " + run_command(job, "goto", url=config.login_url or config.url)["message"])
        elif choice in {"c", "ㅊ"}:
            job.command("capture", phase="after")
            print("  " + wait(job, quiet=True)["message"])
        elif choice in {"q", "ㅂ"}:
            return False


def autofill(job, username, password=None, site=""):
    remembered = password is not None
    if password is None:
        password = getpass.getpass("  비밀번호 (입력 내용은 화면에 표시되지 않음): ")
    if not password:
        print("  비밀번호가 비어 있어 건너뜁니다.")
        return
    ticket = job.command("fill", username=username, password=password)
    print("  " + job.wait_done(ticket)["message"])
    if not remembered and ask_yes(f"  같은 사이트({site})의 다음 대상에도 이 계정을 채울까요? (이번 실행 동안 메모리에만, 다른 사이트에는 채우지 않음)", True):
        Account.username, Account.password, Account.site = username, password, site
    del password


STATUS_HELP = {
    "수집됨": "페이지를 열거나 메뉴를 눌러 확인함",
    "동작 확인 필요": "주소 없이 스크립트로 움직이는 메뉴인데 눌러서 확인하지 못함 (로그아웃·삭제 같은 이름은 일부러 누르지 않음)",
    "미방문": "메뉴 목록에는 있지만 열어 보지 못함 (중지·제한, 또는 링크 깊이 초과)",
    "로그인 필요": "열면 로그인 화면이 나옴. 로그인 후 메뉴도 필요하면 수집 방식 4(전·후 비교)",
    "로그인·권한 확인": "HTTP 401/403: 로그인이나 권한이 있어야 열림",
    "접속 실패": "열리지 않음 (HTTP 오류·시간 초과)",
    "범위 밖": "다른 호스트나 같은 호스트의 다른 서비스라 수집하지 않음",
    "범위 밖 이동": "열었더니 다른 호스트로 넘어감 (로그인 서버 등이면 고급 설정의 추가 호스트)",
    "파일 링크": "다운로드 파일이라 열지 않음",
    "화면 읽기 실패": "열렸지만 메뉴 구조를 읽지 못함",
}
ISSUE_HELP = {
    "이름 없는 링크": ("글자 없이 이미지·아이콘만 있는 링크", "메뉴라면 결과 엑셀 '미수집_확인필요' 시트에서 위치를 확인합니다."),
    "같은 페이지 반복": ("주소의 토큰·시각 값만 바뀐 같은 화면이 반복됨", "서로 다른 화면인데 이렇게 나오면 알려 주세요."),
    "로그인 필요": ("로그인 화면이 감지된 주소", "로그인 후 메뉴가 필요하면 수집 방식 4(전·후 비교)로 다시 수집합니다."),
    "로그인·권한 확인": ("401/403 응답", "권한이 있는 계정으로 로그인 후 수집합니다."),
    "접속 실패": ("열리지 않은 주소", "주소·VPN·사내망을 확인합니다."),
    "범위 밖 이동": ("다른 호스트로 넘어간 주소", "필요하면 고급 설정의 추가 호스트에 넣습니다."),
    "경로 범위 밖": ("같은 호스트의 다른 서비스 주소", "다른 진단 시트의 대상이라 이 시트에는 넣지 않았습니다."),
    "탐색 이동 제한": ("링크 깊이 제한으로 열지 않은 주소", "필요하면 고급 설정의 링크 깊이를 늘립니다."),
    "수집 미완료": ("중지·제한으로 끝까지 확인하지 못함", "다시 수집하거나 고급 설정의 제한을 늘립니다."),
    "동적 메뉴 확인": ("마우스를 올려도 열리지 않은 메뉴", "로그인 단계에서 메뉴를 직접 열고 c(현재 화면 추가)를 누릅니다."),
    "시작 페이지 접속 실패": ("시작 주소가 열리지 않음", "주소·VPN·사내망을 확인합니다."),
    "로그인 상태 확인": ("로그인 후에도 비밀번호 칸이 보임", "로그인이 끝났는지 확인하고 다시 수집합니다."),
    "프레임 범위 밖": ("다른 호스트의 프레임", "필요하면 고급 설정의 추가 호스트에 넣습니다."),
    "화면 읽기 실패": ("화면을 읽지 못함", "메뉴 영역 CSS 선택자를 지정해 다시 수집합니다."),
}


def summary(data):
    title("수집 결과")
    menus = data["menus"]
    ready = sum(m["ready"] for m in menus)
    clicked = sum(1 for m in menus if "눌러" in (m.get("reason") or ""))
    print(f"  발견 메뉴 {len(menus)}개 | 확인된 메뉴 {ready}개 | 방문 페이지 {data['visited_count']}개"
          + (f" | 눌러서 확인 {clicked}개" if clicked else ""))
    vis = Counter(m["visibility"] for m in menus if m.get("visibility"))
    if vis and vis["로그인 후 전용"]:
        print(f"  로그인 전 공개 {vis['로그인 전 공개']}개 | 로그인 후 전용 {vis['로그인 후 전용']}개")
    after = "after" in data.get("phases_done", [])
    status_help = {**STATUS_HELP, **({"로그인 필요": "로그인한 뒤에도 로그인 화면이 나옴 (그 메뉴 권한이 없거나 로그인이 풀림)"} if after else {})}
    issue_help = {**ISSUE_HELP, **({"로그인 필요": ("로그인한 뒤에도 로그인 화면이 나온 주소", "계정 권한을 확인하거나 다시 로그인해 수집합니다.")} if after else {})}
    print()
    print("  상태별 메뉴")
    by_status = {}
    for m in menus:
        if m["kind"] != "group":
            by_status.setdefault(m["status"], []).append(" > ".join(m["path"][:3]))
    for status, names in sorted(by_status.items(), key=lambda x: (x[0] != "수집됨", -len(x[1]))):
        print(f"    - {status} {len(names)}개: " + ui.dim(status_help.get(status, "")))
        if status != "수집됨":
            shown = " | ".join(names[:5]) + (f"  ... 외 {len(names) - 5}개" if len(names) > 5 else "")
            ui.say(ui.dim(shown), indent="        ")
    open_issues = [x for x in data["issues"] if not x.get("resolved")]
    if open_issues:
        print()
        print(f"  확인 필요 {len(open_issues)}건 " + ui.dim("(결과 엑셀 '미수집_확인필요' 시트에 전체 목록)"))
        groups = {}
        for item in open_issues:
            groups.setdefault(item["kind"], []).append(item)
        for kind, items in sorted(groups.items(), key=lambda x: -len(x[1])):
            meaning, todo = issue_help.get(kind, (items[0].get("reason", ""), ""))
            print(f"    - {kind} {len(items)}건: " + ui.dim(meaning))
            if todo:
                print("        " + ui.dim("할 일: " + todo))
            for url in list(dict.fromkeys(x["url"] for x in items if x.get("url")))[:3]:
                ui.say(ui.dim(url), indent="        ")
    if data["limit_reached"]:
        print()
        print("  ⚠ 수집이 끝까지 가지 못했습니다(중지 또는 제한). 필요하면 다시 수집하거나 고급 설정을 늘리세요.")
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
    if out:
        path = Path(out)
    else:
        name = "".join("_" if c in '\\/:*?"<>|' else c for c in data["config"]["site_name"]).strip(" .") or "메뉴수집"
        base = Path("결과") / f"{name}_메뉴목록_{datetime.now():%Y%m%d_%H%M%S}"
        path, n = base.with_name(base.name + ".xlsx"), 2      # not with_suffix: names like www.example.com
        while path.exists():
            path, n = base.with_name(f"{base.name}_{n}.xlsx"), n + 1
    path.parent.mkdir(parents=True, exist_ok=True)
    content = make_workbook(data)
    try:
        path.write_bytes(content)
    except PermissionError:
        # Usually the same file is open in Excel on Windows.
        path = path.with_name(f"{path.stem}_{datetime.now():%H%M%S}.xlsx")
        print(f"  ※ 같은 이름의 파일이 열려 있어 다른 이름으로 저장합니다: {path.name}")
        path.write_bytes(content)
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
    text = text.rstrip("\r\n")  # a trailing newline would paste an empty 4th row
    return win32_clipboard(text) or powershell_clipboard(text)


def win32_clipboard(text):
    """Copy as CF_UNICODETEXT through the Win32 API (no PowerShell needed)."""
    try:
        user32, kernel32 = ctypes.WinDLL("user32"), ctypes.WinDLL("kernel32")
        kernel32.GlobalAlloc.restype = ctypes.c_void_p
        kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalFree.argtypes = [ctypes.c_void_p]
        user32.SetClipboardData.restype = ctypes.c_void_p
        user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        data = text.encode("utf-16-le") + b"\x00\x00"
        for _ in range(10):  # another program may hold the clipboard for a moment
            if user32.OpenClipboard(None):
                break
            time.sleep(0.1)
        else:
            return False
        try:
            user32.EmptyClipboard()
            handle = kernel32.GlobalAlloc(0x0002, len(data))  # GMEM_MOVEABLE
            pointer = kernel32.GlobalLock(handle)
            ctypes.memmove(pointer, data, len(data))
            kernel32.GlobalUnlock(handle)
            if not user32.SetClipboardData(13, handle):  # CF_UNICODETEXT
                kernel32.GlobalFree(handle)
                return False
            return True
        finally:
            user32.CloseClipboard()
    except (OSError, AttributeError):
        return False


def powershell_clipboard(text):
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", suffix=".txt", delete=False) as f:
        f.write(text)
        temp = f.name
    try:
        quoted = temp.replace("'", "''")
        command = f"Get-Content -Raw -Encoding UTF8 -LiteralPath '{quoted}' | Set-Clipboard"
        return subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
                              capture_output=True, timeout=20).returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False
    finally:
        os.unlink(temp)


# ---------- 진단표 기입 ----------

BOARD_KEYS = {"1": "gsheet", "2": "google", "3": "excel", "4": None}


def all_connections():
    found = gsheet.connections()
    return found or [c for c in [gsheet.load_config()] if c]


def connect_google(interactive, force=False):
    """A saved connection, or set one up. --connect-google (force) adds another account's connection."""
    found = all_connections()
    if found and not force:
        return gsheet.Client(found[0]["url"], found[0]["token"], name=found[0].get("name", ""))
    if not interactive:
        sys.exit("구글 시트 연결 정보가 없습니다. 먼저 run.bat --connect-google 로 연결하세요.")
    ui.header("구글 시트 연결", "계정마다 한 번")
    saved = gsheet.file_connections()
    if saved:
        print("  저장된 연결")
        for i, conn in enumerate(saved, 1):
            print(f"    {i}) {conn.get('name', '')}  " + ui.dim(conn["url"][:60] + "..."))
        choice = ask("  Enter = 새 연결 추가 | 번호 앞에 d = 그 연결 삭제 (예: d1)").strip().lower()
        if choice.startswith("d") and choice[1:].isdigit():
            left = gsheet.remove_connection(int(choice[1:]) - 1)
            ui.ok(f"삭제했습니다. 남은 연결 {len(left)}개")
            return None
    print("  1. https://script.google.com -> 새 프로젝트 (그 계정으로 로그인한 상태에서, 계정마다 한 번만)")
    print(f"  2. {Path(__file__).resolve().parent.parent / 'apps_script' / 'MenuImport.gs'} 내용을 전부 붙여넣고 저장")
    print("  3. 함수 setup 실행 -> 권한 승인 -> 실행 로그의 '연결 키' 복사 (또는 프로젝트 설정 -> 스크립트 속성 TOKEN)")
    print("  4. 배포 -> 새 배포 -> 웹 앱 (실행: 나 / 액세스: 모든 사용자) -> 웹 앱 URL 복사")
    print("  " + ui.dim("진단표 구글 시트 주소는 실행할 때마다 묻고, 그 시트를 열 수 있는 연결을 자동으로 씁니다."))
    while True:
        url = ask("  웹 앱 URL (취소하려면 Enter)")
        if not url:
            return None
        token = ask("  연결 키")
        try:
            client = gsheet.Client(url, token)
            title_ = verify_connection(client)
        except SheetError as exc:
            ui.fail(str(exc))
            continue
        name = ask("  이 연결의 이름 (예: 개인 계정, 회사 계정)", f"연결 {len(saved) + 1}")
        path = gsheet.save_config(client.url, client.token, name)
        client.name = name
        ui.ok(f"연결했습니다: {name} ({title_})  " + ui.dim(f"연결 정보: {path.name}, 이 PC에만 저장"))
        return client


def pick_connection(ref, interactive):
    """The one connection (Google account) used for this sheet: remembered per sheet, asked the first time.
    Other accounts are never tried on a sheet, so a sheet is only written through the account chosen for it."""
    conns = all_connections()
    remembered = state.sheet_connection(ref)
    hit = next((c for c in conns if c["url"] == remembered), None)
    if hit:
        return hit
    if len(conns) == 1 and not remembered:
        return conns[0]
    if not interactive:
        sys.exit("이 시트에 쓸 구글 연결을 정하지 않았습니다. 한 번 대화형으로 실행해 연결을 고르세요.")
    if remembered:
        ui.note("이 시트를 쓰던 연결이 더 이상 없습니다(삭제됨). 이 시트에 쓸 연결을 다시 고르세요.")
    print("  이 시트에 쓸 구글 연결(계정)을 고르세요. " + ui.dim("고른 연결로만 이 시트에 씁니다. 다음부터는 묻지 않습니다."))
    for i, conn in enumerate(conns, 1):
        print(f"    {i}) {conn.get('name', '') or f'연결 {i}'}")
    while True:
        pick = ask("  번호")
        if pick.isdigit() and 1 <= int(pick) <= len(conns):
            state.remember_sheet_connection(ref, conns[int(pick) - 1]["url"])
            return conns[int(pick) - 1]
        print(f"  1~{len(conns)} 중에서 고르세요.")


def open_sheet(client, ref, interactive):
    conn = pick_connection(ref, interactive)
    client = gsheet.Client(conn["url"], conn["token"], spreadsheet=ref, name=conn.get("name", ""))
    print("  구글 시트를 읽는 중입니다..." + (ui.dim(f"  (연결: {client.name})") if client.name else ""))
    board = SheetsBoard(client)
    try:
        board.cached = board.targets()
    except SheetError as exc:
        ui.fail(f"{client.name or '연결'}: {exc}")
        if exc.transport:      # network / deployment trouble: says nothing about the sheet, keep its connection
            ui.note("네트워크나 웹 앱 배포 문제입니다. 연결은 그대로 두었습니다. 잠시 후 다시 실행하세요.")
            sys.exit(1)
        ui.note("이 연결의 계정으로는 이 시트를 열 수 없습니다. 시트 주소와 그 계정의 편집 권한을 확인하세요.")
        return None
    if client.outdated():
        ui.note(f"배포된 Apps Script가 예전 버전입니다 (배포본 {client.version or '버전 없음'} / 최신 {gsheet.SCRIPT_VERSION}). "
                "붙여넣고 저장만 하면 예전 코드가 계속 실행됩니다.")
        print("  " + gsheet.UPDATE_STEPS)
        print("  " + ui.dim("예전 버전에는 최근 안전장치(진단 시트 확인, 쓰기 직전 재확인, 정리 점검)가 없습니다."))
        if interactive and not ask_yes("  예전 버전 그대로 계속할까요?", False):
            sys.exit("Apps Script를 새 버전으로 배포한 뒤 다시 실행하세요.")
    return board


def verify_connection(client):
    """The web app answers and the key is right (a sheet address is not needed yet)."""
    try:
        return client.ping()
    except SheetError as exc:
        if "주소가 필요" in str(exc):          # standalone deployment: fine, the address comes at run time
            return "연결 확인"
        raise


def choose_board(args, interactive):
    """Which 진단표 the menus go into. None = collect only."""
    kind = args.target or ("excel" if args.workbook else None)
    path = args.workbook
    if not kind:
        if not interactive:
            kind = "google"
        else:
            linked = bool(all_connections())
            last = "" if linked else state.last_workbook()
            if linked:
                last_sheet, last_title = state.last_sheet()
                print("  " + ui.bold("구글 시트 주소 붙여넣기" + (f" 또는 Enter(지난번 시트 '{last_title or '진단표'}')" if last_sheet else " 또는 Enter"))
                      + ui.dim("  = 구글 시트에 바로 입력, 값만 입력"))
                print("  " + ui.dim("엑셀 파일로 하려면 파일을 끌어다 놓기 | p = 구글 시트 붙여넣기 | - = 기입 없이 수집만 | c = 정리 점검"))
            else:
                print("  작업할 진단 엑셀 파일을 붙여넣어 주세요. (파일을 이 창에 끌어다 놓아도 됩니다)")
                if last:
                    print("  " + ui.dim(f"Enter = 지난번 파일 {Path(last).name}"))
                print("  " + ui.dim("g = 구글 시트에 바로 입력 (처음 1회 연결) | p = 구글 시트 붙여넣기 | "
                                    + ("- = 기입 없이 수집만" if last else "Enter = 기입 없이 수집만")))
            while True:
                value = ask("  진단표" if linked else "  작업 엑셀").strip().strip('"').strip("'").strip()
                if gsheet.sheet_ref(value):            # a Google Sheets address pasted right here
                    args.sheet_url, value = value, "g"
                if value.lower() in {"c", "ㅊ"}:
                    args.check = True
                    print("  " + ui.bold("정리 점검") + ui.dim(": 이미 들어간 메뉴 열 중 잘못 들어간 열을 찾습니다. 이어서 진단표를 고르세요."))
                    continue
                if not value and linked:
                    value = "g"
                if not value and last:
                    value = last
                if value.lower() in {"g", "ㅎ"}:
                    kind = "gsheet"
                elif value.lower() in {"p", "ㅔ"}:
                    kind = "google"
                elif not value or value == "-":
                    return None
                elif Path(value).is_file() and value.lower().endswith((".xlsx", ".xlsm")):
                    kind, path = "excel", value
                else:
                    print("  구글 시트 주소(https://docs.google.com/spreadsheets/...)나 .xlsx/.xlsm 파일 경로를 넣어 주세요.")
                    continue
                break
    if kind == "gsheet":
        try:
            client = connect_google(interactive)
        except SheetError as exc:
            sys.exit(str(exc))
        if client is None:
            if interactive and not args.target:
                return choose_board(args, interactive)
            return None
        last, last_title = state.last_sheet()
        ref = gsheet.sheet_ref(args.sheet_url) if args.sheet_url else ""
        while not ref:
            if not interactive:
                ref = last
                if not ref:
                    sys.exit("--sheet-url 로 진단표 구글 시트 주소를 지정하세요.")
                break
            hint = f" (Enter = 지난번 시트 '{last_title or last}')" if last else ""
            raw = ask(f"  진단표 구글 시트 주소{hint} (- = 처음 질문으로)")
            if raw.strip() == "-":
                args.target, args.sheet_url = None, ""
                return choose_board(args, interactive)
            ref = gsheet.sheet_ref(raw) if raw else last
            if not ref:
                print("  구글 시트 주소를 붙여넣으세요. 예: https://docs.google.com/spreadsheets/d/.../edit")
        board = open_sheet(client, ref, interactive)
        if board is None:
            if interactive:
                args.sheet_url = ""
                return choose_board(args, interactive)
            sys.exit("구글 시트를 열 수 없습니다.")
        state.remember_sheet(ref, board.title)
        ui.ok(f"구글 시트 '{board.title}' | 진단 시트 {len(board.cached)}개")
        return board
    if kind == "google":
        snapshot = args.snapshot
        if not snapshot and interactive:
            print("  " + ui.dim("이미 메뉴가 있는 시트에 이어 붙이려면 구글 시트를 .xlsx로 내려받아 지정합니다. (파일 → 다운로드 → Microsoft Excel)"))
            snapshot = ask_path("  내려받은 .xlsx 경로 (빈 시트에 붙여넣을 거면 Enter)")
        return PasteBoard(snapshot, copy_to_clipboard if os.name == "nt" else None)
    if kind == "excel":
        path = path or (ask_path("  작업할 진단 엑셀 파일을 붙여넣어 주세요") if interactive else "")
        if not path:
            return None
        try:
            board = ExcelBoard(path, args.workbook_out)
            count = len(board.targets())
        except Exception as exc:  # noqa: BLE001 - any unreadable workbook: say so and ask again
            if not interactive:
                sys.exit(f"진단 엑셀을 읽을 수 없습니다: {exc}")
            ui.fail(f"진단 엑셀을 읽을 수 없습니다: {exc}")
            args.workbook = ""
            return choose_board(args, interactive)
        if not count:
            ui.fail("이 파일에서 진단 시트(01_, 02_ … 형식)를 찾지 못했습니다.")
            if not interactive:
                sys.exit(1)
            args.workbook = ""
            return choose_board(args, interactive)
        ui.ok(f"{Path(path).name} | 진단 시트 {count}개")
        ui.say(ui.dim(f"기입 결과는 새 파일 {board.out.name} 에 저장합니다. 원본은 그대로 둡니다."))
        return board
    return None


def show_targets(items):
    print("  " + ui.dim("진단 시트 목록  (괄호 = 이미 입력된 메뉴 수)"))
    ui.grid([t.sheet + (f" ({t.menus})" if t.menus else "") for t in items])


def load_targets(board):
    try:
        return board.targets()
    except SheetError as exc:
        sys.exit(str(exc))


def map_url(url, items, interactive, query="", stop_if_missing=True):
    """Inspection URL -> 진단 시트 by 대표 도메인 (D6). -> (sheet or None, sheets to choose from).
    Clear fits are taken; a fit by a different path or a parent domain is shown and confirmed."""
    if query:
        return None, items
    level, found = diag.match_levels(url, items)
    reason = diag.MATCH_REASON.get(level, "")
    if len(found) == 1 and level >= 3:
        ui.ok(f"{reason}: {found[0].sheet}")
        return found[0], items
    if len(found) == 1 and interactive:
        ui.note(f"{reason}: {found[0].sheet}  (D6: {found[0].addresses[0]})")
        if ask_yes(f"  이 주소를 {found[0].sheet}에 기입할까요?", True):
            return found[0], items
        return None, items
    if len(found) > 1:
        ui.note(f"{reason} 후보가 여러 개입니다. 하나를 고르세요.")
    elif level:
        pass
    else:
        ui.note("대표 도메인(D6)이 맞는 진단 시트가 없습니다. 진단 대상을 입력하세요.")
    if not interactive:
        if stop_if_missing:
            sys.exit("--sheet 로 진단 시트를 지정하세요. (진단 시트 D6에 대표 도메인을 적어 두면 자동으로 찾습니다)")
        return None, items
    return None, found or items


def choose_targets(board, args, interactive, many=False, url="", items=None):
    items = load_targets(board) if items is None else items
    query = args.sheet
    pool = items
    if url and items:
        hit, pool = map_url(url, items, interactive, query)
        if hit:
            return [hit]
    if not items:   # paste mode without a downloaded copy: sheet names as typed
        if not query and interactive:
            query = ask("  붙여넣을 진단 시트 이름" + (" (여러 개는 쉼표로 구분)" if many else ""))
        names = [n.strip() for n in (query or "").split(",") if n.strip()] or [""]
        return [diag.Target(n, n) for n in (names if many else names[:1])]
    if interactive and not query and (pool is not items or not board.listed):
        show_targets(pool)     # the full list once per run; a narrowed list every time
        board.listed = board.listed or pool is items
    while True:
        if not query:
            if not interactive:
                sys.exit("--sheet 로 진단 시트를 지정하세요. 예: --sheet 01_사이트A")
            query = ask(f"  진단 대상 (이름 일부만 써도 됨, {'여러 개는 쉼표, ' if many else ''}목록 다시 보기 ?)")
            if not query:
                print("  진단 대상을 입력하세요.")
                continue
            if query in {"?", "？"}:
                show_targets(pool)
                query = ""
                continue
        try:
            picked = diag.select_targets(query, pool)
            if picked and (many or len(picked) == 1):
                if not many and url:
                    ui.ok(picked[0].sheet)
                return picked
            ui.fail("이 단계에서는 진단 대상을 하나만 지정하세요.")
        except ValueError as exc:
            if not interactive:
                sys.exit(str(exc))
            for msg in str(exc).splitlines():
                ui.fail(msg)
        query = ""


def fill_step(menus, args, interactive, out_dir, url=""):
    """One crawl result -> one 진단 시트 (used after a single crawl and with --from-result)."""
    if getattr(args, "collect_only", False):
        return None
    if interactive:
        ui.header("진단표 기입")
    board = choose_board(args, interactive)
    if board is None:
        return None
    target = choose_targets(board, args, interactive, url=url)[0]
    print()
    print("  " + ui.bold(f"> {target.sheet or '진단 시트'}"))
    try:
        return board.fill(target, menus, interactive, out_dir)
    except Exception as exc:  # noqa: BLE001 - report instead of a traceback; the crawl result is already saved
        ui.fail(f"기입하지 못했습니다: {type(exc).__name__}: {exc}")
        return Filled(target, "실패")


def parse_args(argv=None):
    p = argparse.ArgumentParser(prog="메뉴수집", description="웹사이트 메뉴를 1·2·3 Depth로 수집해 엑셀로 정리합니다. 옵션 없이 실행하면 하나씩 물어봅니다.")
    p.add_argument("url", nargs="?", help="수집할 사이트 주소")
    p.add_argument("--mode", choices=["auto", "public", "login", "compare"],
                   help="auto=필요할 때만 로그인(기본), public=로그인 전, login=로그인 후, compare=전·후 비교")
    p.add_argument("--login-url", default="", help="로그인 페이지 주소")
    p.add_argument("--name", default="", help="대상 이름")
    p.add_argument("--out", default="", help="저장할 엑셀 파일 경로 (기본: 결과 폴더)")
    p.add_argument("--username", help="로그인 단계에서 자동 입력할 아이디 (비밀번호는 실행 중에 따로 물어봄)")
    p.add_argument("--max-pages", type=int, default=0, help="최대 방문 페이지 수 (0 = 제한 없음, 기본)")
    p.add_argument("--max-seconds", type=int, default=0, help="대상별 최대 수집 시간(초) (0 = 제한 없음, 기본)")
    p.add_argument("--max-hops", type=int, default=6)
    p.add_argument("--wait-ms", type=int, default=700)
    p.add_argument("--delay-ms", type=int, default=300)
    p.add_argument("--extra-host", action="append", default=[], help="추가 허용 호스트 (여러 번 지정 가능)")
    p.add_argument("--selector", default="", help="메뉴 영역 CSS 선택자")
    p.add_argument("--browser", choices=["auto", "msedge", "chrome", "chromium"], default="auto")
    p.add_argument("--headless", action="store_true", help="브라우저 창 없이 실행 (로그인 전 수집만)")
    p.add_argument("--no-click", action="store_true", help="javascript: 메뉴를 눌러 보지 않음")
    p.add_argument("--my-browser", action="store_true", help="내 브라우저 창에서 로그인 (인증서·OTP·보안 프로그램이 필요한 사이트)")
    p.add_argument("--check", action="store_true", help="정리 점검: 이미 들어간 메뉴 열 중 잘못 들어간 열을 찾아 비움")
    p.add_argument("-y", "--yes", action="store_true", help="확인 질문 없이 바로 실행 (점검 권한 확인 완료로 간주)")
    p.add_argument("--no-open", action="store_true", help="저장 후 엑셀을 자동으로 열지 않음")
    p.add_argument("--workbook", default="", help="메뉴를 입력할 진단 엑셀(.xlsx) 경로")
    p.add_argument("--sheet", default="", help="메뉴를 입력할 진단 시트 이름 (예: 01_사이트A)")
    p.add_argument("--workbook-out", default="", help="메뉴를 입력한 진단 엑셀을 저장할 경로 (기본: 원본 옆에 새 파일)")
    p.add_argument("--target", choices=["gsheet", "google", "excel"], help="진단표 기입 방법: gsheet=구글 시트 자동 입력, google=구글 시트 붙여넣기, excel=엑셀 파일 직접 입력")
    p.add_argument("--connect-google", action="store_true", help="구글 시트 자동 입력 연결(웹 앱 URL·연결 키) 설정")
    p.add_argument("--sheet-url", default="", help="진단표 구글 시트 주소 (구글 시트 자동 입력, 기본: 지난번 시트)")
    p.add_argument("--snapshot", default="", help="구글 시트를 엑셀로 내려받은 파일 (붙여넣을 위치·중복 계산용)")
    p.add_argument("--include-unverified", action="store_true", help="로그인 필요·접속 실패 등으로 확인하지 못한 메뉴도 진단 시트에 입력")
    p.add_argument("--from-result", default="", help="크롤링 없이, 이전에 만든 수집 결과 엑셀의 메뉴를 진단 엑셀에 입력")
    return p.parse_args(argv)



# ---------- 실행 ----------

def ask_path(prompt):
    while True:
        value = ask(prompt).strip().strip('"').strip("'")
        if not value:
            return ""
        if Path(value).is_file() and value.lower().endswith((".xlsx", ".xlsm")):
            return value
        print("  .xlsx 또는 .xlsm 파일 경로를 확인하세요. (파일을 이 창으로 끌어다 놓아도 됩니다)")



def choose_menus(confirmed, unverified, args, interactive):
    """confirmed + (optionally) navigation menus that could not be confirmed, e.g. behind a login. Tidied for the sheet."""
    if not unverified:
        return diag.tidy(confirmed)
    if interactive and getattr(args, "unverified_asked", False):
        print("  " + ui.dim(f"확인하지 못한 메뉴 {len(unverified)}개: " + ("포함" if args.include_unverified else "제외") + " (첫 대상에서 고른 대로)"))
    if not args.include_unverified and interactive and not getattr(args, "unverified_asked", False):
        args.unverified_asked = True
        from collections import Counter
        kinds = " · ".join(f"{k} {n}" for k, n in Counter(s for _, s in unverified).most_common())
        print()
        print(f"  메뉴 구조에서 이름은 확인했지만 페이지를 확인하지 못한 메뉴 {len(unverified)}개 ({kinds})")
        if any(k == "주소 없는 메뉴" for _, k in unverified):
            print("  " + ui.dim("주소 없는 메뉴 = 클릭하면 스크립트로 화면을 바꾸는 메뉴(javascript:). 진단 시트에는 메뉴 이름만 들어갑니다."))
        args.include_unverified = ask_yes("  이 메뉴도 진단 시트에 넣을까요? (이후 대상에도 같은 답을 적용)", True)
    return diag.tidy(confirmed + [d for d, _ in unverified] if args.include_unverified else confirmed)




def ask_login_window(args):
    if args.my_browser:
        return
    print()
    print("  로그인할 창")
    print("    1) 수집기 창 (기본)   " + ui.dim("아이디/비밀번호, 문자·앱 OTP"))
    print("    2) 내 브라우저 창     " + ui.dim("공동·금융인증서, 키보드 보안 등 보안 프로그램, 자동화 창을 막는 사이트"))
    args.my_browser = ask("  번호", "2" if state.pref("my_browser") else "1") == "2"
    state.remember("my_browser", args.my_browser)
    if args.my_browser:
        print("  " + ui.dim("Edge/Chrome을 수집기 전용 프로필(browser_profile)로 엽니다. 보안 프로그램·확장은 그 창에서 한 번만 설치하면 됩니다."))


def engine_mode(mode):
    """'auto' runs as before/after comparison; the login part is skipped when the public crawl is enough."""
    return "compare" if mode == "auto" else mode


def needs_login(data):
    """Auto mode: log in when the public crawl shows (almost) only a login wall.
    - every page failed to open: logging in would not help (unreachable, VPN off) -> False
    - no confirmed menu besides the login functions -> True
    - most opened pages (at least two) were login screens -> True"""
    pages = data.get("pages", [])
    if pages and all(p.get("status") in {"접속 실패", "화면 읽기 실패"} for p in pages):
        return False
    confirmed = [m for m in diag.tidy(diag.menus_from_report(data)[0]) if not diag.LOGIN_NAME.match(m[0])]
    walls = sum(1 for p in pages if p.get("status") in {"로그인 필요", "로그인·권한 확인", "범위 밖 이동"})
    login_seen = walls > 0 or any(m.get("evidence") == "로그인 화면" for m in data["menus"])
    if not confirmed:
        return login_seen or not data["menus"]
    return walls >= max(2, 0.6 * len(pages))


def run_crawl(config, args, auto=False):
    """Open the browser, crawl (with the login step when needed) and close it. -> (data or None, error or None)."""
    print("  브라우저를 여는 중입니다...")
    job = Job(config)
    try:
        data = wait(job, quiet=True)
        print("  " + data["message"])
        if job.stop.is_set():          # q while the browser was opening: do not start an unlimited crawl
            print("  중지 요청으로 수집을 시작하지 않습니다.")
            return None, KeyboardInterrupt()
        if config.mode in {"public", "compare"}:
            if config.mode == "compare":
                print("\n  [1단계] 로그인 전 공개 메뉴 수집")
            data = crawl(job, "before")
        if auto and config.mode == "compare" and not job.stop.is_set():
            if not needs_login(data):
                print("  공개 메뉴가 보여 로그인 없이 마칩니다. " + ui.dim("(로그인 후 메뉴까지 필요하면 수집 방식 4)"))
                return data, None
            print("  " + ui.bold("로그인 화면만 보입니다. 로그인한 뒤 메뉴를 수집합니다."))
        if config.mode in {"login", "compare"} and not job.stop.is_set():
            step = login_step(job, config, args.username)
            if step == "public":
                if config.mode == "compare":
                    print("  로그인 전 수집 결과만 남깁니다.")
                    data = job.snapshot()
                else:
                    print("\n  [로그인 없이] 공개 메뉴 수집")
                    data = crawl(job, "before")
            elif step:
                ui.window_title("로그인 후 수집 중")
                print("\n  [로그인 후] 메뉴 수집")
                data = crawl(job, "after")
                for _ in range(3):              # the session ended on the way: log in again and go on
                    if not data.get("session_lost") or job.stop.is_set() or args.yes:
                        break
                    print()
                    ui.note("수집 중에 로그인이 풀렸습니다(연속으로 로그인 화면). 지금까지 수집한 메뉴는 남아 있습니다.")
                    if not ask_yes("  다시 로그인하고 이어서 수집할까요?", True):
                        break
                    if login_step(job, config, args.username) is not True:
                        break
                    ui.window_title("로그인 후 수집 중")
                    print("\n  [로그인 후] 남은 페이지부터 이어서 수집")
                    data = crawl(job, "after", resume=True)
            else:
                data = job.snapshot()
        return data, (KeyboardInterrupt() if job.stop.is_set() else None)
    except (KeyboardInterrupt, EngineError) as exc:
        print(f"\n  실행 오류: {exc}" if isinstance(exc, EngineError) else "\n  중단했습니다.")
        job.stop.set()
        data = job.snapshot()
        if data["menus"]:
            print("  지금까지 수집한 결과를 저장합니다.")
            return data, exc
        return None, exc
    finally:
        job.close()
        job.thread.join(timeout=10)


def finish(data, args):
    summary(data)
    path = save(data, args.out)
    print()
    ui.ok(f"메뉴 목록 엑셀: {path}")
    confirmed, unverified = diag.menus_from_report(data, include_unverified=True)
    fill_step(choose_menus(confirmed, unverified, args, not args.yes), args, not args.yes, Path(path).parent, data["config"]["url"])
    if os.name == "nt" and not args.no_open and not args.yes and ask_yes("메뉴 목록 엑셀 파일을 열까요?", False):
        os.startfile(path)  # noqa: S606 - Windows only, local result file


def single_run(args):
    config = build_config(args)
    title(MODE_LABELS["auto" if getattr(args, "auto_login", False) else config.mode])
    data, error = run_crawl(config, args, auto=getattr(args, "auto_login", False))
    if data is None:
        if isinstance(error, EngineError):
            sys.exit(1)
        return
    if error is not None:
        args.yes = True  # no more questions after an interruption
    finish(data, args)


def campaign(args):
    """작업 엑셀 → 점검 주소(여러 개, 대표 도메인으로 시트 매핑) → 수집 방식 → 대상마다 수집·저장·기입 → 결과 표."""
    ui.header("웹 메뉴 수집기", "수집 -> 메뉴 목록 엑셀 -> 진단 시트 기입")
    ui.step(1, 3, "진단표")
    board = choose_board(args, True)
    if args.check and board is not None:
        return check_board(board, args)
    if board is None:
        args.collect_only = True
        return single_run(args)
    items = load_targets(board)
    ui.step(2, 3, "진단 대상")
    plan = pick_plan(board, args, items)
    ui.step(3, 3, "수집 방식 (모든 대상 공통)")
    mode = ask_mode(args.mode, default="1")
    if mode != "public":
        ask_login_window(args)
        print()
        print("  " + ui.dim("로그인 페이지는 보통 점검 주소에서 바로 이동할 수 있습니다."))
        if ask_yes("  점검 주소와 다른 로그인 페이지 주소를 따로 지정할 대상이 있나요?", False):
            plan = [(t, url, ask_optional_url(f"    {t.sheet or '진단 시트'} 로그인 주소 (없으면 Enter)")) for t, url, _ in plan]
    settings = base_settings(args)
    advanced_settings(args, settings, True)

    ui.header("작업 순서 확인", board.label)
    print("  접속 확인 중...")
    reach = netcheck.check_all([url for _, url, _ in plan])
    ui.table([("순서", 4), ("진단 시트", 28), ("기존 메뉴", 9), ("접속", 18), ("점검 주소", 0)],
             [[str(i), t.sheet or "진단 시트", str(t.menus) if t.menus else "-", ("" if ok else "! ") + text, url]
              for i, ((t, url, _), (ok, text)) in enumerate(zip(plan, reach), 1)])
    if not all(ok for ok, _ in reach):
        ui.note("접속되지 않는 대상이 있습니다. VPN, 사내망, 주소를 확인하세요. 그대로 시작하면 해당 대상은 '수집 실패'로 남습니다.")
    if any(t.menus for t, _, _ in plan):
        print("  " + ui.dim("기존 메뉴가 있는 시트는 이미 있는 메뉴를 건너뛰고 새 메뉴만 이어서 넣습니다."))
    print()
    print(f"  수집 방식: {MODE_LABELS[mode]} | {limits(settings['max_pages'], settings['max_seconds'])}")
    print("  " + ui.dim("메뉴를 모두 확인하면 자동으로 끝납니다. 중간에 멈추려면 q (그때까지 결과는 저장)"))
    if mode != "public":
        print("  " + ui.dim("대상마다 브라우저 창이 열리면 직접 로그인한 뒤 이 창에서 Enter를 누릅니다."))
    if any(t.save_domain for t, _, _ in plan):
        print("  " + ui.dim("대표 도메인(D6)이 비어 있는 시트는 점검 주소도 함께 적습니다."))
    if not ask_yes("  모든 대상의 점검 권한(계약, 승인)을 확인했고, 이대로 시작할까요?", True):
        sys.exit("취소했습니다.")

    done, stop = run_plan(plan, items, mode, settings, board, args)
    retry = [(row, entry) for row, entry in zip(done, plan) if row[0].status in RETRY]
    report(done, board, mode, args, stop, offer_open=not retry or stop is not None)
    while retry and stop is None:
        print()
        names = ", ".join(row[0].target.sheet or "진단 시트" for row, _ in retry)
        if not ask_yes(f"  끝나지 않은 대상 {len(retry)}개(실패·건너뜀·실행 안 함·일부 입력: {names[:70]})를 지금 다시 시도할까요?", False):
            break
        # crawled fine but not filled (sheet error, 'no' at the fill question): fill again from the saved menu list
        refill = [(row, entry) for row, entry in retry if row[0].status in REFILL and row[2] and Path(row[2]).is_file()]
        recrawl = [entry for row, entry in retry if (row, entry) not in refill]
        done = [refill_row(row, board, args) for row, _ in refill]
        if recrawl:
            more, stop = run_plan(recrawl, items, mode, settings, board, args)
            done += more
        report(done, board, mode, args, stop)
        entries = [entry for _, entry in refill] + recrawl
        retry = [(row, entry) for row, entry in zip(done, entries) if row[0].status in RETRY]
    if isinstance(stop, SystemExit):
        raise stop


RETRY = {"수집 실패", "실패", "중단", "건너뜀", "실행 안 함", "기입할 메뉴 없음", "기입 안 함", "일부 입력", "수집 미완료"}


REFILL = {"실패", "기입 안 함", "건너뜀"}


def refill_row(row, board, args):
    """Fill again from the menu-list Excel saved by the crawl (no new crawl, no new login)."""
    filled, _, path, url = row
    t = filled.target
    ui.header(f"다시 기입: {t.sheet}", Path(path).name)
    try:
        menus = choose_menus(diag.menus_from_result(path), diag.unverified_from_result(path), args, True)
        return board.fill(t, menus, True, Path(path).parent), len(menus), path, url
    except (ValueError, OSError, SheetError) as exc:
        ui.fail(f"{t.sheet}: {exc}")
        return Filled(t, "실패"), 0, path, url


def run_plan(plan, items, mode, settings, board, args):
    """Crawl and fill the planned targets in order. -> (rows for the result table, stop exception or None)."""
    done, idx, current, stop = [], 0, None, None
    started = time.monotonic()
    try:
        while idx < len(plan):
            t, url, login_url = plan[idx]
            if idx:
                print()
                choice = ask(f"  다음 대상 {t.sheet or '진단 시트'} | Enter=시작 · s=이 대상 건너뛰기 · q=여기서 끝내기").strip().lower()
                if choice in {"q", "ㅂ"}:
                    break
                if choice in {"s", "ㄴ"}:
                    done.append((Filled(t, "건너뜀"), 0, "", url))
                    idx += 1
                    continue
            current = (t, url)
            ui.window_title(f"[{idx + 1}/{len(plan)}] {t.sheet} 수집 중")
            try:
                scope = diag.crawl_scope(t, url, items)
                if scope:
                    print("  " + ui.dim(f"같은 호스트에 다른 진단 대상이 있어 {scope} 아래만 수집합니다."))
                done.append(run_target(idx + 1, len(plan), t, url, login_url, mode, {**settings, "path_prefix": scope}, board, args))
            except (ValueError, SheetError) as exc:     # e.g. another sheet layout: a clear message, not a crash
                ui.fail(f"{t.sheet or '진단 시트'}: {exc}")
                done.append((Filled(t, "실패"), 0, "", url))
            except Exception as exc:  # noqa: BLE001 - one broken target must not stop the others
                ui.fail(f"{t.sheet or '진단 시트'} 처리 중 오류: {type(exc).__name__}: {exc}")
                done.append((Filled(t, "실패"), 0, "", url))
            current = None
            idx += 1
    except (KeyboardInterrupt, SystemExit) as exc:
        stop = exc
        print()
        ui.note("중단했습니다. 지금까지 한 작업을 정리합니다.")
        if current:
            done.append((Filled(current[0], "중단"), 0, "", current[1]))
            idx += 1
    done += [(Filled(t, "실행 안 함"), 0, "", url) for t, url, _ in plan[idx:]]
    minutes, seconds = divmod(int(time.monotonic() - started), 60)
    print()
    print("  " + ui.dim(f"걸린 시간: {minutes}분 {seconds}초 · 대상 {len(plan)}개"))
    ui.window_title("작업 끝")
    ui.beep()
    return done, stop


def report(done, board, mode, args, stop=None, offer_open=True):
    """Result table, work log, where the filled workbook is."""
    ui.header("작업 결과", board.label)
    ui.table([("진단 시트", 26), ("수집", 5), ("기입", 17), ("중복", 5), ("위치", 13), ("점검 주소", 0)],
             [[f.target.sheet or "진단 시트", str(n) if n else "-",
               f.status + (f" {f.written}개" if f.written else "") + (f" (칸 부족 {f.overflow})" if f.overflow else ""),
               str(f.duplicates) if f.duplicates else "-", f.where or "-", url] for f, n, _, url in done])
    short = sum(f.overflow for f, _, _, _ in done)
    if short:
        ui.note(f"빈 열이 모자라 넣지 못한 메뉴 {short}개가 있습니다. 입력 내역 파일(.tsv)의 '칸 부족' 줄을 확인하세요.")
    files = [name for _, _, name, _ in done if name]
    if files:
        print("  " + ui.dim(f"메뉴 목록 엑셀 {len(files)}개: {(Path('결과')).resolve()}"))
    log = write_log(done, board, mode)
    if log:
        print("  " + ui.dim(f"작업 기록: {log}"))
    print()
    if isinstance(board, ExcelBoard) and board.current == board.src:
        ui.note("새로 기입한 메뉴가 없어 새 진단 엑셀을 만들지 않았습니다. 작업 엑셀은 그대로입니다.")
    if isinstance(board, ExcelBoard):
        state.remember_workbook(board.current)       # next run: Enter continues from the latest copy
    if isinstance(board, ExcelBoard) and board.current != board.src:
        ui.ok("진단 엑셀 (원본은 그대로):")
        print("  " + str(board.out.resolve()))
        if os.name == "nt" and offer_open and not args.no_open and not args.yes and stop is None and ask_yes("  진단 엑셀을 열까요?", True):
            os.startfile(board.out)  # noqa: S606 - Windows only, local file the user just created
    if isinstance(board, SheetsBoard):
        if any(f.status == "입력 완료" for f, _, _, _ in done):
            ui.ok(f"구글 시트 '{board.title}':")
            print("  " + board.client.spreadsheet)
            for f, _, _, _ in done:
                if f.status == "입력 완료" and f.file.startswith("http"):
                    print("    " + ui.dim(f"{f.target.sheet}: ") + f.file)
            if os.name == "nt" and offer_open and not args.no_open and not args.yes and stop is None and ask_yes("  구글 시트를 브라우저로 열까요?", True):
                import webbrowser
                webbrowser.open(board.client.spreadsheet)
        else:
            ui.note("구글 시트에 새로 입력한 메뉴가 없습니다.")
    if any(f.status == "붙여넣기 준비" for f, _, _, _ in done):
        print("  " + ui.dim("붙여넣기 파일과 입력 내역(.tsv)은 결과 폴더에 있습니다."))


def result_files(spec):
    """--from-result: a file, a folder (every *_메뉴목록_*.xlsx in it) or a wildcard pattern."""
    import glob
    path = Path(spec)
    if path.is_dir():
        found = sorted(path.glob("*_메뉴목록_*.xlsx"))
    elif any(ch in spec for ch in "*?["):
        found = sorted(Path(p) for p in glob.glob(spec))
    else:
        found = [path] if path.is_file() else []
    return [p for p in found if p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")]


def fill_results(files, args):
    """Several crawl result files -> their 진단 시트 (mapped by each file's start address and D6), one output copy."""
    interactive = not args.yes
    ui.header("메뉴 목록 엑셀 -> 진단 시트 기입", f"결과 파일 {len(files)}개")
    board = choose_board(args, interactive)
    if board is None:
        return
    items = load_targets(board)
    done = []
    try:
        for f in files:
            url = diag.result_url(f)
            print()
            print("  " + ui.bold(f"> {f.name}") + ui.dim(f"  {url}"))
            try:
                menus = diag.menus_from_result(f)
                unverified = diag.unverified_from_result(f)
            except (ValueError, OSError) as exc:
                ui.fail(f"읽을 수 없는 파일이라 건너뜁니다: {exc}")
                continue
            hit, pool = map_url(url, items, interactive, stop_if_missing=False) if url else (None, items)
            if hit is None:
                if not interactive:
                    done.append((Filled(diag.Target(""), "시트 못 찾음"), len(menus), f.name, url))
                    continue
                hit = pick_one(pool, items, "  진단 대상 (이름 일부 · ?=목록 · Enter=이 파일 건너뛰기)")
                if hit is None:
                    done.append((Filled(diag.Target(""), "건너뜀"), len(menus), f.name, url))
                    continue
            menus = choose_menus(menus, unverified, args, interactive)
            if not menus:
                ui.fail("넣을 메뉴가 없습니다.")
                done.append((Filled(hit, "기입할 메뉴 없음"), 0, f.name, url))
                continue
            try:
                done.append((board.fill(hit, menus, interactive, f.parent), len(menus), f.name, url))
            except Exception as exc:  # noqa: BLE001 - one broken file must not stop the others
                ui.fail(f"기입하지 못했습니다: {type(exc).__name__}: {exc}")
                done.append((Filled(hit, "실패"), len(menus), f.name, url))
    except KeyboardInterrupt:
        print()
        ui.note("중단했습니다. 지금까지 한 작업을 정리합니다.")
    report(done, board, "", args)


def check_board(board, args):
    """정리 점검: list wrong menu columns already on the sheets, and empty them only after the user agrees."""
    if isinstance(board, PasteBoard):
        sys.exit("정리 점검은 구글 시트 연결 또는 엑셀 파일에서만 됩니다.")
    items = [t for t in load_targets(board) if t.menus]
    ui.header("정리 점검", f"{board.label}")
    if not items:
        ui.ok("메뉴가 들어간 진단 시트가 없습니다.")
        return
    print(f"  메뉴가 들어간 진단 시트 {len(items)}개")
    ui.grid([f"{t.sheet} ({t.menus})" for t in items])
    raw = ask("  점검할 시트 (Enter = 모두, 번호·이름·범위도 됨)")
    try:
        picked = diag.select_targets(raw, items) if raw else items
    except ValueError as exc:
        sys.exit(str(exc))
    total = 0
    for t in picked:
        print()
        print("  " + ui.bold(f"> {t.sheet}"))
        try:
            found = diag.audit(board.columns(t.sheet))
        except (ValueError, SheetError) as exc:
            ui.fail(str(exc))
            continue
        if not found:
            ui.ok("이상 없음")
            continue
        ui.table([("열", 4), ("이유", 30), ("결과", 6), ("메뉴", 0)],
                 [[diag.L(f["col"]), f["reason"], f"{f['results']}개" if f["results"] else "-",
                   " > ".join(x for x in f["menu"] if x)] for f in found], indent="    ")
        safe = [f for f in found if not f["results"]]
        if len(safe) < len(found):
            ui.note(f"진단 결과가 있는 열 {len(found) - len(safe)}개는 비우지 않습니다. 확인 후 직접 정리하세요.")
        if not safe or not ask_yes(f"  위 열 중 {len(safe)}개의 메뉴 이름(13~15행)을 비울까요?", False):
            continue
        res = board.clear(t.sheet, safe)
        total += len(res["cleared"])
        ui.ok(f"{len(res['cleared'])}개 열을 비웠습니다: {', '.join(diag.L(c) for c in res['cleared'])}")
        if res["skipped"]:
            ui.note(f"그사이 바뀌었거나 결과가 생긴 열 {len(res['skipped'])}개는 건너뛰었습니다.")
    print()
    if isinstance(board, ExcelBoard) and board.current != board.src:
        ui.ok(f"정리한 진단 엑셀 (원본은 그대로): {board.out.resolve()}")
    ui.ok(f"정리 점검을 마쳤습니다. 비운 열 {total}개")


def write_log(done, board, mode):
    """결과/작업기록_날짜.tsv: one line per target, for the report and for picking up where the run stopped."""
    if not done:
        return None
    workbook = str(board.out.resolve()) if isinstance(board, ExcelBoard) and board.current != board.src else ""
    if isinstance(board, SheetsBoard):
        workbook = f"구글 시트 {board.title}: {board.client.spreadsheet}"
    lines = ["\t".join(["진단 시트", "점검 주소", "수집 방식", "수집 메뉴", "기입 결과", "입력", "중복 제외", "칸 부족", "위치",
                        "메뉴 목록 엑셀", "진단표"])]
    for f, n, name, url in done:
        cells = [f.target.sheet, url, MODE_LABELS.get(mode, mode), n, f.status, f.written, f.duplicates, f.overflow, f.where, name,
                 workbook if f.status == "입력 완료" else ""]
        lines.append("\t".join(" ".join(str(x).split()) for x in cells))
    path = Path("결과") / f"작업기록_{datetime.now():%Y%m%d_%H%M%S}.tsv"
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8-sig")
    except OSError:
        return None
    return path.resolve()


def split_tokens(raw):
    """'29-34, 36 41' -> ['29-34', '36', '41']: commas always split; spaces split only lists of numbers."""
    out = []
    for token in [x.strip() for x in re.split(r"[,，]", raw) if x.strip()]:
        if re.fullmatch(r"[\d\s~-]+", token):
            out += re.sub(r"\s*[-~]\s*", "-", token).split()
        else:
            out.append(token)
    return out


def exact_sheet(token, items):
    key = "".join(token.split()).lower()
    hits = [t for t in items if key in {"".join(t.sheet.split()).lower(), "".join(t.name.split()).lower()}]
    return hits[0] if len(hits) == 1 else None


def ask_optional_url(prompt):
    """An address, or '' when the user presses Enter (no forced loop after a typo)."""
    while True:
        raw = ask(prompt)
        if not raw:
            return ""
        if diag.address(raw):
            return ask_url(raw)
        print("  주소 형식이 아닙니다. 예: www.example.com  (건너뛰려면 Enter)")


def pick_one(pool, items, prompt):
    """One sheet out of pool (or any sheet by exact name); '' = skip."""
    while True:
        query = ask(prompt)
        if not query:
            return None
        if query.strip() in {"?", "？"}:
            show_targets(pool)
            continue
        for scope in (pool, items) if pool is not items else (items,):
            try:
                found = diag.select_targets(query, scope)
            except ValueError:
                continue
            if len(found) == 1:
                ui.ok(found[0].sheet)
                return found[0]
        ui.fail(f"'{query}'에 해당하는 진단 시트를 하나로 정할 수 없습니다. 이름을 더 적거나 ?로 목록을 보세요.")


def pick_plan(board, args, items):
    """Sheets by number/name/'전체' (their D6 addresses are used) and/or inspection URLs (mapped by D6).
    -> [(target, url, login_url)] without duplicates; one entry per address."""
    with_domain = [t for t in items if t.addresses]
    if with_domain:
        print(f"  대표 도메인(D6)이 적힌 시트 {len(with_domain)}개 " + ui.dim("(번호나 이름만 입력하면 D6 주소로 수집)"))
        ui.grid([t.sheet for t in with_domain])
    print()
    if items:
        print("  진단할 시트 번호나 이름, 또는 점검 주소를 입력하세요. 여러 개는 쉼표로 구분합니다.")
        print("  " + ui.dim("예: 01-05, 07, 12 | 사이트A | 전체 (D6이 적힌 시트 모두) | https://... | ? = 전체 시트 목록"))
    else:
        print("  붙여넣을 진단 시트 이름이나 점검 주소를 입력하세요. 여러 개는 쉼표로 구분합니다.")
    plan = []
    key = lambda name: "".join((name or "").split()).lower()  # noqa: E731

    def add(t, url):
        if not t.layout:
            ui.fail(f"{t.sheet}: 13~15행이 1·2·3 Depth 행이 아니라 이 시트에는 기입할 수 없습니다(양식 확인). 대상에서 뺐습니다.")
            return
        if any(key(x.sheet) == key(t.sheet) and u == url for x, u, _ in plan):
            ui.note(f"이미 넣은 대상이라 건너뜁니다: {t.sheet}  {url}")
            return
        again = any(key(x.sheet) == key(t.sheet) for x, _, _ in plan)
        plan.append((t, url, ""))
        print(f"    + {len(plan)}. {t.sheet or '진단 시트'}  {url}" + ui.dim("  (같은 시트에 이어서 기입)" if again else ""))

    def offer_d6(t, url):
        if board.can_save_domain and t.sheet and not t.d6 and not diag.match_url(url, [t]):
            if ask_yes(f"    이 주소를 {t.sheet} 대표 도메인(D6)에 적어 둘까요? 다음부터 번호만으로 수집합니다", True):
                t.save_domain = url

    def by_url(url):
        if not items:
            name = ask("    이 주소를 붙여넣을 진단 시트 이름 (빼려면 Enter)")
            return add(diag.Target(name, name), url) if name else ui.note("이 주소는 뺐습니다.")
        hit, pool = map_url(url, items, True)
        if hit is None:
            if pool is items:
                if not board.listed:
                    show_targets(items)
                board.listed = True
            else:
                ui.grid([t.sheet for t in pool])
            hit = pick_one(pool, items, "  진단 대상 (이름 일부 · ?=목록 · Enter=이 주소 빼기)")
            if hit is None:
                return ui.note("이 주소는 뺐습니다.")
            offer_d6(hit, url)
        add(hit, url)

    def by_sheet(t):
        urls = t.addresses
        if not urls:
            url = ask_optional_url(f"    {t.sheet} 점검 주소 (D6에 주소가 없음, 건너뛰려면 Enter)")
            if not url:
                return ui.note(f"{t.sheet} 은(는) 주소가 없어 뺐습니다.")
            offer_d6(t, url)
            urls = [url]
        if len(urls) > 1:
            print("    " + ui.dim(f"{t.sheet}: D6 주소 {len(urls)}개를 차례로 수집해 같은 시트에 기입합니다."))
        for url in urls:
            add(t, ask_url(url))

    last_targets = state.pref("targets") if items else ""
    while True:
        raw = ask("  대상" + (f" (Enter = 지난번 {last_targets})" if last_targets and not plan else "")
                  if not plan else "  추가 대상 (다 입력했으면 Enter)")
        if not raw and not plan and last_targets:
            raw, last_targets = last_targets, ""
        if not raw:
            if plan:
                return plan
            print("  대상을 하나 이상 입력하세요.")
            continue
        if not re.match(r"https?://", raw.strip(), re.I):
            state.remember("targets", raw.strip())
        if raw.strip() in {"?", "？"}:
            show_targets(items)
            continue
        for token in split_tokens(raw):
            if re.match(r"https?://", token, re.I) or (diag.address(token) and not exact_sheet(token, items)):
                if diag.address(token):
                    by_url(ask_url(token))
                else:
                    ui.fail(f"'{token}'은(는) 주소 형식이 아닙니다.")
                continue
            if not items:                    # paste mode without a downloaded copy: the name is the sheet
                by_sheet(diag.Target(token, token))
                continue
            try:
                picked = with_domain if token in {"전체", "all", "*"} else diag.select_targets(token, items)
            except ValueError as exc:
                ui.fail(str(exc))
                continue
            if not picked:
                ui.fail("대표 도메인(D6)이 적힌 시트가 없습니다. 시트 번호나 주소로 입력하세요.")
            for t in picked:
                by_sheet(t)


def run_target(i, total, t, url, login_url, mode, settings, board, args):
    """One 진단 대상: crawl -> save the menu list -> fill the 진단 시트. -> (Filled, menus, result file, url)."""
    ui.header(f"[{i}/{total}] {t.name or t.sheet}", f"{t.sheet}  {url}")
    if args.my_browser:
        try:
            settings = {**settings, "cdp_url": mybrowser.start(url, headless=args.headless)}
        except RuntimeError as exc:
            ui.fail(str(exc))
            return Filled(t, "실패"), 0, "", url
    try:
        config = CrawlConfig(url=url, mode=engine_mode(mode), site_name=(t.name or t.sheet or urlsplit(url).hostname or "메뉴 수집")[:120],
                             login_url=login_url, extra_hosts=args.extra_host, browser=args.browser,
                             menu_selector=args.selector, headless=args.headless, **settings)
    except ValidationError as exc:
        ui.fail("입력값 오류: " + "; ".join(e["msg"] for e in exc.errors()))
        return Filled(t, "실패"), 0, "", url
    data, error = run_crawl(config, args, auto=mode == "auto")
    if data is None:
        ui.fail("수집한 메뉴가 없습니다.")
        return Filled(t, "수집 실패"), 0, "", url
    summary(data)
    path = save(data, "")
    print()
    ui.ok(f"메뉴 목록 엑셀: {path.name}")
    print("  " + ui.dim(str(path.parent)))
    if not data["menus"]:
        why = next((p.get("reason") or p.get("status") for p in data.get("pages", []) if p.get("status") != "수집됨"), "")
        ui.fail("메뉴를 수집하지 못해 기입하지 않았습니다." + (f" 사유: {why}" if why else " 주소와 접속(VPN, 사내망)을 확인하세요."))
        return Filled(t, "수집 실패"), 0, str(path), url
    confirmed, unverified = diag.menus_from_report(data, include_unverified=True)
    menus = choose_menus(confirmed, unverified, args, True)
    if not menus:
        counts = ", ".join(f"{k} {n}" for k, n in Counter(m["status"] for m in data["menus"]).most_common(4))
        ui.fail(f"진단 시트에 넣을 메뉴가 없어 기입하지 않았습니다. (발견 {len(data['menus'])}개: {counts})")
        ui.say(ui.dim("메뉴가 표준 형태가 아니면 고급 설정의 '메뉴 영역 CSS 선택자'나 로그인 단계의 c(현재 화면 추가)를 사용하세요."))
        return Filled(t, "기입할 메뉴 없음"), 0, str(path), url
    if error is not None and not ask_yes(f"  수집을 중간에 멈췄습니다. 멈춘 시점까지의 메뉴 {len(menus)}개를 진단 시트에 기입할까요?", True):
        return Filled(t, "기입 안 함"), len(menus), str(path), url
    print()
    print("  " + ui.bold(f"> 진단 시트 기입: {t.sheet or '진단 시트'}"))
    filled = board.fill(t, menus, True, path.parent)
    if (data.get("session_lost") or data.get("limit_reached")) and filled.status in {"입력 완료", "새 메뉴 없음"}:
        filled.status = "일부 입력" if filled.status == "입력 완료" else "수집 미완료"
        ui.note("수집이 끝까지 가지 못한 대상이라 결과 표에 '" + filled.status + "'로 남기고, 다시 시도 목록에 넣습니다.")
    return filled, len(menus), str(path), url


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    args = parse_args(argv)
    ui.install_console()
    if args.connect_google:
        connect_google(True, force=True)
        return
    if args.from_result:
        files = result_files(args.from_result)
        if not files:
            sys.exit(f"수집 결과 파일을 찾을 수 없습니다: {args.from_result}")
        if len(files) == 1:
            title("수집 결과 -> 진단 시트 입력")
            try:
                menus = diag.menus_from_result(files[0])
                unverified = diag.unverified_from_result(files[0])
            except (ValueError, OSError) as exc:
                sys.exit(f"수집 결과 파일을 읽을 수 없습니다: {exc}")
            print(f"  수집 결과에서 진단 메뉴 {len(menus)}개를 읽었습니다.")
            menus = choose_menus(menus, unverified, args, not args.yes)
            fill_step(menus, args, not args.yes, files[0].parent, diag.result_url(files[0]))
            return
        return fill_results(files, args)
    try:
        if not args.url and not args.yes:
            return campaign(args)
        single_run(args)
    except KeyboardInterrupt:
        print()
        sys.exit("  중단했습니다.")


if __name__ == "__main__":
    main()
