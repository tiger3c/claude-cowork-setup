"""Google Sheets 진단표 connection through the Apps Script web app in apps_script/MenuImport.gs.

The web app writes values only into rows 13-15 of a 진단 시트 (after the last used column) and never
changes formats, formulas or other cells. The connection (web app URL + 연결 키) is kept on this PC only,
in google_sheet.json next to run.bat (ignored by git).
"""
import json
from urllib.parse import urlsplit
import re
import os
import urllib.error
import urllib.request
from pathlib import Path

from . import workbook as diag

CONFIG = Path(__file__).resolve().parent.parent / "google_sheet.json"


class SheetError(Exception):
    """transport=True: the web app could not be reached or answered with HTTP/HTML (network, deployment),
    so nothing is known about the sheet itself; False: the web app answered (e.g. no access to that sheet)."""

    def __init__(self, message, transport=False):
        super().__init__(message)
        self.transport = transport


def _env():
    url, token = os.environ.get("MENU_CRAWLER_GSHEET_URL", ""), os.environ.get("MENU_CRAWLER_GSHEET_TOKEN", "")
    return {"name": "환경 변수", "url": url, "token": token} if url and token else None


def file_connections():
    """Saved connections, one per Google account that deployed the web app (e.g. personal, company)."""
    try:
        data = json.loads(CONFIG.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    if isinstance(data, dict) and isinstance(data.get("connections"), list):
        return [c for c in data["connections"] if isinstance(c, dict) and c.get("url") and c.get("token")]
    if isinstance(data, dict) and data.get("url") and data.get("token"):          # older single-connection file
        return [{"name": "연결 1", "url": data["url"], "token": data["token"]}]
    return []


def connections():
    env = _env()
    return [env] if env else file_connections()


def load_config():
    found = connections()
    return found[0] if found else None


def save_config(url, token, name=""):
    """Add a connection (a connection with the same web app URL is replaced)."""
    keep = [c for c in file_connections() if c["url"] != url]
    keep.append({"name": name or f"연결 {len(keep) + 1}", "url": url, "token": token})
    CONFIG.write_text(json.dumps({"connections": keep}, ensure_ascii=False, indent=1), encoding="utf-8")
    return CONFIG


def remove_connection(index):
    keep = file_connections()
    if 0 <= index < len(keep):
        gone = keep.pop(index)
        from . import state
        state.forget_connection(gone["url"])          # sheets bound to it must be chosen again
        CONFIG.write_text(json.dumps({"connections": keep}, ensure_ascii=False, indent=1), encoding="utf-8")
    return keep


SHEET_URL = re.compile(r"^(?:https?://)?docs\.google\.com/spreadsheets/(?:u/\d+/)?d/([a-zA-Z0-9_-]{20,})")


def sheet_ref(text):
    """A Google Sheets address (or its id) -> the canonical address; '' when it is not one."""
    text = (text or "").strip().strip('"').strip("'")
    m = SHEET_URL.match(text)
    if m:
        return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/edit"
    return f"https://docs.google.com/spreadsheets/d/{text}/edit" if re.fullmatch(r"[a-zA-Z0-9_-]{30,}", text) else ""


class Client:
    def __init__(self, url, token, timeout=120, spreadsheet="", name=""):
        url = url.strip()
        local_test = os.environ.get("MENU_CRAWLER_GSHEET_TEST") == "1" and urlsplit(url).scheme == "http" \
            and urlsplit(url).hostname == "127.0.0.1" and "@" not in urlsplit(url).netloc
        if not url.startswith("https://script.google.com/") and not local_test:
            raise SheetError("웹 앱 URL은 https://script.google.com/ 으로 시작해야 합니다. (배포 → 웹 앱 URL)")
        self.url, self.token, self.timeout, self.spreadsheet, self.name = url, token.strip(), timeout, spreadsheet, name

    def call(self, action, **payload):
        body = json.dumps({"action": action, "token": self.token, "spreadsheet": self.spreadsheet, **payload},
                          ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(self.url, data=body, headers={"Content-Type": "application/json"}, method="POST")
        try:
            # Apps Script answers POST with a redirect to the result; urllib follows it with GET as required.
            with urllib.request.urlopen(req, timeout=self.timeout) as res:
                text = res.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as exc:
            raise SheetError(f"구글 시트 연결 실패 (HTTP {exc.code}). 웹 앱 배포의 액세스 권한이 '모든 사용자'인지 확인하세요.",
                             transport=True) from exc
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            raise SheetError(f"구글 시트에 연결할 수 없습니다(네트워크): {getattr(exc, 'reason', exc)}", transport=True) from exc
        try:
            data = json.loads(text)
        except ValueError as exc:
            hint = "로그인 화면이 돌아왔습니다. 웹 앱 액세스 권한을 '모든 사용자'로 배포하세요." if "<html" in text[:500].lower() else text[:120]
            raise SheetError(f"구글 시트 응답을 읽을 수 없습니다. {hint}", transport=True) from exc
        if not data.get("ok"):
            raise SheetError(data.get("error") or "구글 시트에서 오류가 발생했습니다.")
        return data

    def ping(self):
        return self.call("ping")["title"]

    def targets(self):
        data = self.call("targets")
        return data["title"], [diag.Target(t["sheet"], t.get("name") or "", diag._domain(t.get("domain", "")),
                                           int(t.get("menus", 0)), int(t.get("start", diag.FIRST_COL)), diag.clean(t.get("domain", ""), 2000),
                                           layout=t.get("layout", True) is not False)
                               for t in data["targets"]]

    def fill(self, sheet, menus, dry_run=False, domain=""):
        rows = [[diag.clean(x) for x in (list(m) + ["", "", ""])[:3]] for m in menus]
        return self.call("fill", sheet=sheet, menus=[r for r in rows if r[0]], dryRun=dry_run, domain=diag.clean(domain))

    def columns(self, sheet):
        return self.call("columns", sheet=sheet)["columns"]

    def clear(self, sheet, items):
        return self.call("clear", sheet=sheet, columns=[{"col": i["col"], "menu": i["menu"]} for i in items])
