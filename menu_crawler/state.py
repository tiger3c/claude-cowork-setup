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
