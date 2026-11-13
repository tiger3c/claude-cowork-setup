"""The user's own browser window for sites that need certificates, OTP or security programs to log in.

Edge (or Chrome) is started normally, not as an automation browser, with a profile kept next to run.bat
(browser_profile/) so security programs and extensions are installed once. The crawler attaches to it through
the local DevTools port; only the crawler's own tab is used and closed.
"""
import os
import shutil
import subprocess
import time
import urllib.request
from pathlib import Path

PORT = int(os.environ.get("MENU_CRAWLER_CDP_PORT", "9222"))
PROFILE = Path(__file__).resolve().parent.parent / "browser_profile"


def endpoint(port=None):
    return f"http://127.0.0.1:{port or PORT}"


def running(port=None):
    try:
        with urllib.request.urlopen(endpoint(port) + "/json/version", timeout=1):
            return True
    except OSError:
        return False


def find_browser():
    custom = os.environ.get("MENU_CRAWLER_BROWSER_PATH", "")
    if custom:
        return custom
    if os.name == "nt":
        for base in (os.environ.get("ProgramFiles(x86)", ""), os.environ.get("ProgramFiles", ""), os.environ.get("LOCALAPPDATA", "")):
            for rel in (r"Microsoft\Edge\Application\msedge.exe", r"Google\Chrome\Application\chrome.exe"):
                path = Path(base) / rel
                if base and path.is_file():
                    return str(path)
    for name in ("msedge", "microsoft-edge", "google-chrome", "chrome", "chromium"):
        found = shutil.which(name)
        if found:
            return found
    return ""


def start(url, port=None, headless=False):
    """Open (or reuse) the user's browser window and return its DevTools endpoint."""
    port = port or PORT
    if running(port):
        return endpoint(port)
    exe = find_browser()
    if not exe:
        raise RuntimeError("Edge 또는 Chrome을 찾지 못했습니다.")
    PROFILE.mkdir(exist_ok=True)
    args = [exe, f"--remote-debugging-port={port}", f"--user-data-dir={PROFILE}", "--no-first-run",
            "--no-default-browser-check", "--disable-features=Translate,TranslateUI", url]
    if headless:
        args[1:1] = ["--headless=new"]
    if os.name != "nt" and hasattr(os, "geteuid") and os.geteuid() == 0:
        args[1:1] = ["--no-sandbox"]            # Chromium refuses to start as root otherwise (Linux test machines)
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=flags)
    for _ in range(60):
        if running(port):
            return endpoint(port)
        time.sleep(0.25)
    raise RuntimeError("내 브라우저 창을 열었지만 연결할 수 없습니다. 열린 창을 모두 닫고 다시 시도하세요.")
