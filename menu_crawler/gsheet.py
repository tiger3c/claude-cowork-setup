"""Google Sheets 진단표 connection through the Apps Script web app in apps_script/MenuImport.gs.

The web app writes values only into rows 13-15 of a 진단 시트 (after the last used column) and never
changes formats, formulas or other cells. The connection (web app URL + 연결 키) is kept on this PC only,
in google_sheet.json next to run.bat (ignored by git).
"""
import json
import os
import urllib.error
import urllib.request
from pathlib import Path

from . import workbook as diag

CONFIG = Path(__file__).resolve().parent.parent / "google_sheet.json"


class SheetError(Exception):
    pass


def load_config():
    url, token = os.environ.get("MENU_CRAWLER_GSHEET_URL", ""), os.environ.get("MENU_CRAWLER_GSHEET_TOKEN", "")
    if url and token:
        return {"url": url, "token": token}
    try:
        data = json.loads(CONFIG.read_text(encoding="utf-8"))
        return data if data.get("url") and data.get("token") else None
    except (OSError, ValueError):
        return None


def save_config(url, token):
    CONFIG.write_text(json.dumps({"url": url, "token": token}, ensure_ascii=False, indent=1), encoding="utf-8")
    return CONFIG


class Client:
    def __init__(self, url, token, timeout=120):
        if not url.startswith("https://script.google.com/"):
            raise SheetError("웹 앱 URL은 https://script.google.com/ 으로 시작해야 합니다. (배포 → 웹 앱 URL)")
        self.url, self.token, self.timeout = url.strip(), token.strip(), timeout

    def call(self, action, **payload):
        body = json.dumps({"action": action, "token": self.token, **payload}, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(self.url, data=body, headers={"Content-Type": "application/json"}, method="POST")
        try:
            # Apps Script answers POST with a redirect to the result; urllib follows it with GET as required.
            with urllib.request.urlopen(req, timeout=self.timeout) as res:
                text = res.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as exc:
            raise SheetError(f"구글 시트 연결 실패 (HTTP {exc.code}). 웹 앱 배포의 액세스 권한이 '모든 사용자'인지 확인하세요.") from exc
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            raise SheetError(f"구글 시트에 연결할 수 없습니다: {getattr(exc, 'reason', exc)}") from exc
        try:
            data = json.loads(text)
        except ValueError as exc:
            hint = "로그인 화면이 돌아왔습니다. 웹 앱 액세스 권한을 '모든 사용자'로 배포하세요." if "<html" in text[:500].lower() else text[:120]
            raise SheetError(f"구글 시트 응답을 읽을 수 없습니다. {hint}") from exc
        if not data.get("ok"):
            raise SheetError(data.get("error") or "구글 시트에서 오류가 발생했습니다.")
        return data

    def ping(self):
        return self.call("ping")["title"]

    def targets(self):
        data = self.call("targets")
        return data["title"], [diag.Target(t["sheet"], t.get("name") or "", diag._domain(t.get("domain", "")),
                                           int(t.get("menus", 0)), int(t.get("start", diag.FIRST_COL)), diag.clean(t.get("domain", "")))
                               for t in data["targets"]]

    def fill(self, sheet, menus, dry_run=False, domain=""):
        rows = [[diag.clean(x) for x in (list(m) + ["", "", ""])[:3]] for m in menus]
        return self.call("fill", sheet=sheet, menus=[r for r in rows if r[0]], dryRun=dry_run, domain=diag.clean(domain))
