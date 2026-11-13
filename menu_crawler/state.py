"""Small per-PC conveniences (the last work Excel). Kept next to run.bat in last_run.json, ignored by git."""
import json
import os
from pathlib import Path

FILE = Path(__file__).resolve().parent.parent / "last_run.json"


def _load():
    if os.environ.get("MENU_CRAWLER_NO_STATE"):
        return {}
    try:
        return json.loads(FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def last_workbook():
    path = _load().get("workbook", "")
    return path if path and Path(path).is_file() else ""


def remember_workbook(path):
    if os.environ.get("MENU_CRAWLER_NO_STATE"):
        return
    try:
        data = _load()
        data["workbook"] = str(Path(path).resolve())
        FILE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass


def last_sheet():
    data = _load()
    return data.get("sheet", ""), data.get("sheet_title", "")


def remember_sheet(url, title=""):
    if os.environ.get("MENU_CRAWLER_NO_STATE"):
        return
    try:
        data = _load()
        data["sheet"], data["sheet_title"] = url, title
        FILE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass


def sheet_connection(sheet):
    return _load().get("sheet_connections", {}).get(sheet, "")


def remember_sheet_connection(sheet, connection_url):
    if os.environ.get("MENU_CRAWLER_NO_STATE"):
        return
    try:
        data = _load()
        links = data.setdefault("sheet_connections", {})
        if connection_url:
            links[sheet] = connection_url
        else:
            links.pop(sheet, None)
        FILE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass


def pref(key, default=""):
    return _load().get("prefs", {}).get(key, default)


def remember(key, value):
    if os.environ.get("MENU_CRAWLER_NO_STATE"):
        return
    try:
        data = _load()
        data.setdefault("prefs", {})[key] = value
        FILE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass


def forget_connection(connection_url):
    if os.environ.get("MENU_CRAWLER_NO_STATE"):
        return
    try:
        data = _load()
        links = data.get("sheet_connections", {})
        data["sheet_connections"] = {k: v for k, v in links.items() if v != connection_url}
        FILE.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    except OSError:
        pass
