"""진단표 (the diagnosis board) the crawled menus go into, in three flavours:

- SheetsBoard: the Google Sheet itself, through the Apps Script web app (values only, automatic).
- PasteBoard : Google Sheets by hand - a paste block on the clipboard plus the exact cell to click.
- ExcelBoard : an .xlsx/.xlsm file - a new copy with the menus filled in (the original is untouched).

All of them only ever put menu names into rows 13-15 of columns that hold no menu or result yet.
"""
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from . import ui
from . import workbook as diag
from .gsheet import SheetError


@dataclass
class Filled:
    target: diag.Target
    status: str                 # 입력 완료 / 붙여넣기 준비 / 새 메뉴 없음 / 건너뜀 / 실패
    where: str = ""
    written: int = 0
    duplicates: int = 0
    overflow: int = 0
    file: str = ""


def _range(first, last):
    return f"{diag.L(first)}13:{diag.L(last)}15"


def _preview(menus, limit=5):
    for depths in menus[:limit]:
        print("    - " + " > ".join(d for d in depths if d))
    if len(menus) > limit:
        print("    " + ui.dim(f"… 외 {len(menus) - limit}개"))


def _safe(name):
    return "".join("_" if c in '\\/:*?"<>|' else c for c in name).strip(" .") or "진단시트"


class SheetsBoard:
    kind = "gsheet"
    listed = False
    can_save_domain = True

    def __init__(self, client, title=""):
        self.client, self.title = client, title
        self.label = f"구글 시트 자동 입력 · {title}" if title else "구글 시트 자동 입력"

    cached = None

    def targets(self):
        if self.cached is not None:          # read once when the board was chosen; fills re-read the live sheet
            items, self.cached = self.cached, None
            return items
        self.title, items = self.client.targets()
        self.label = f"구글 시트 자동 입력 · {self.title}" + (f" ({self.client.name})" if self.client.name else "")
        return items

    def columns(self, sheet):
        return self.client.columns(sheet)

    def clear(self, sheet, items):
        return self.client.clear(sheet, items)

    def fill(self, target, menus, interactive, out_dir):
        try:
            pre = self.client.fill(target.sheet, menus, dry_run=True)
            count = len(pre["write"])
            print(f"  시트에 있는 메뉴 {pre['existing']}개 · 새로 입력 {count}개 · 이미 있어 뺀 메뉴 {len(pre['duplicates'])}개"
                  + (f" · 칸 부족 {len(pre['overflow'])}개" if pre["overflow"] else ""))
            if not count:
                ui.ok("새로 넣을 메뉴가 없습니다." + (" 수집한 메뉴가 모두 시트에 이미 있습니다." if pre["duplicates"] else ""))
                if target.save_domain and self.client.fill(target.sheet, [], domain=target.save_domain).get("domainSaved"):
                    ui.ok("대표 도메인(D6)을 적었습니다.")
                return Filled(target, "새 메뉴 없음", duplicates=len(pre["duplicates"]), overflow=len(pre["overflow"]))
            print(f"  입력 위치: ‘{target.sheet}’ {pre['range']}  (값만 입력 · 서식·수식·결과 칸은 그대로)")
            _preview(pre["write"])
            if interactive and not ui.ask_yes("  구글 시트에 바로 입력할까요?", True):
                return Filled(target, "건너뜀")
            res = self.client.fill(target.sheet, menus, domain=target.save_domain)
        except SheetError as exc:
            ui.fail(str(exc))
            return Filled(target, "실패")
        ui.ok(f"구글 시트 ‘{target.sheet}’ {res['range']}에 {res['written']}개 입력했습니다."
              + (" 대표 도메인(D6)도 적었습니다." if res.get("domainSaved") else ""))
        if res.get("url"):
            print("  " + ui.dim("바로 보기: " + res["url"]))
        record = diag.Plan(target.sheet, [], res["existing"], list(zip(range(res["start"], res["start"] + res["written"]), res["write"])),
                           res["duplicates"], res["overflow"])
        rec = diag.receipt(record, Path(out_dir) / f"구글시트입력_{_safe(target.sheet)}_{datetime.now():%Y%m%d_%H%M%S}_입력내역.tsv")
        return Filled(target, "입력 완료", res["range"], res["written"], len(res["duplicates"]), len(res["overflow"]),
                      res.get("url") or str(rec))


class PasteBoard:
    kind = "google"
    listed = False
    can_save_domain = False

    def __init__(self, snapshot="", copy=None):
        self.snapshot, self.copy = snapshot, copy
        self.pasted = {}        # sheet -> {"start", "keys"}: blocks handed out in this run (the snapshot does not know them)
        self.seq = 0
        self.label = "구글 시트 붙여넣기" + (f" · {Path(snapshot).name} 기준" if snapshot else "")

    def targets(self):
        return diag.targets(self.snapshot) if self.snapshot else []

    def fill(self, target, menus, interactive, out_dir):
        sheet = target.sheet
        p = diag.plan_paste(self.snapshot or None, sheet, menus, self.pasted.get(sheet))
        if p.checked:
            print(f"  시트에 있는 메뉴 {p.existing}개 · 새로 붙여넣을 메뉴 {len(p.menus)}개 · 이미 있어 뺀 메뉴 {len(p.duplicates)}개")
        if not p.menus:
            print("  붙여넣을 새 메뉴가 없습니다.")
            return Filled(target, "새 메뉴 없음", duplicates=len(p.duplicates), overflow=len(p.overflow))
        if not p.checked:
            print(f"  붙여넣을 메뉴 {len(p.menus)}개")
            ui.note("메뉴가 없는 빈 진단 시트에만 붙여넣으세요.")
            ui.say(ui.dim("메뉴가 있는 시트라면 구글 시트를 .xlsx로 내려받아 지정하면 이어 붙일 칸을 계산합니다."))
        where = _range(p.start, p.start + len(p.menus) - 1)
        cell = f"{diag.L(p.start)}13"
        self.seq += 1
        base = Path(out_dir) / f"구글시트붙여넣기_{_safe(sheet)}_{datetime.now():%Y%m%d_%H%M%S}_{self.seq}"
        prev = self.pasted.get(sheet, {"keys": set()})
        self.pasted[sheet] = {"start": p.start + len(p.menus), "keys": prev["keys"] | {tuple(m) for m in p.menus}}
        tsv = base.with_suffix(".tsv")
        block = diag.paste_text(p.menus)
        tsv.write_text(block + "\r\n", encoding="utf-8")
        record = diag.Plan(sheet, [], p.existing, list(zip(range(p.start, p.start + len(p.menus)), p.menus)), p.duplicates, p.overflow)
        diag.receipt(record, base.with_name(base.name + "_입력내역.tsv"))
        print()
        print("  " + ui.bold(f"‘{sheet or '진단 시트'}’ 시트의 {cell} 칸을 클릭 → Ctrl+V") + ui.dim(f"   (채워지는 범위 {where})"))
        print("  " + ui.dim(f"붙여넣기 파일: {tsv.resolve()}"))
        if self.copy and self.copy(block):
            ui.ok("클립보드에 복사했습니다.")
        else:
            ui.note("자동 복사를 하지 못했습니다. 붙여넣기 파일을 메모장으로 열어 전체 복사하세요.")
        if p.overflow:
            ui.note(f"칸이 부족해 넣지 못한 메뉴 {len(p.overflow)}개는 입력 내역 파일에 '칸 부족'으로 남겼습니다.")
        if interactive:
            ui.ask("  구글 시트에 붙여넣었으면 Enter")
        return Filled(target, "붙여넣기 준비", where, len(p.menus), len(p.duplicates), len(p.overflow), str(tsv))


class ExcelBoard:
    kind = "excel"
    listed = False
    can_save_domain = True

    def __init__(self, path, out=""):
        self.src = Path(path)
        stem = re.sub(r"(_메뉴입력_\d{8}_\d{6})+$", "", self.src.stem)      # additional runs: no stacked suffixes
        self.out = Path(out) if out else self.src.with_name(f"{stem}_메뉴입력_{datetime.now():%Y%m%d_%H%M%S}{self.src.suffix}")
        if self.out.suffix.lower() != self.src.suffix.lower():      # .xlsm content must keep .xlsm
            self.out = self.out.with_suffix(self.src.suffix)
        if self.out.resolve() == self.src.resolve() or (self.out.exists() and os.path.samefile(self.out, self.src)):
            raise ValueError("결과 파일을 원본과 같은 파일로 지정할 수 없습니다. 원본은 그대로 두고 새 파일로 저장합니다.")
        self.current = self.src
        self.label = f"엑셀 파일 · {self.src.name} → {self.out.name}"

    def targets(self):
        return diag.targets(self.current)

    def fill(self, target, menus, interactive, out_dir):
        p = diag.plan(self.current, target.sheet, menus)
        if target.save_domain and not diag.cell_values(self.current, target.sheet, {"D6"}).get("D6"):
            p.cells["D6"] = target.save_domain
        print(f"  시트에 있는 메뉴 {p.existing}개 · 새로 입력 {len(p.to_write)}개 · 이미 있어 뺀 메뉴 {len(p.duplicates)}개"
              + (f" · 칸 부족 {len(p.overflow)}개" if p.overflow else ""))
        where = _range(p.to_write[0][0], p.to_write[-1][0]) if p.to_write else ""
        if p.to_write:
            print(f"  입력 위치: ‘{target.sheet}’ {where}  (값만 입력 · 서식·수식·결과 칸은 그대로)")
            _preview([d for _, d in p.to_write])
        else:
            ui.ok("새로 넣을 메뉴가 없습니다." + (" 수집한 메뉴가 모두 시트에 이미 있습니다." if p.duplicates else "")
                  + (" 빈 열이 없어 넣지 못한 메뉴는 입력 내역에 남깁니다." if p.overflow else ""))
        if p.to_write or p.cells:
            if not self._save(p, interactive):
                return Filled(target, "실패")
            self.current = self.out
        if p.to_write or p.overflow:
            folder = self.out.parent if p.to_write else Path(out_dir)
            diag.receipt(p, folder / f"{self.out.stem}_{_safe(target.sheet)}_{datetime.now():%H%M%S}_입력내역.tsv")
        if not p.to_write:
            if p.cells:
                ui.ok("대표 도메인(D6)을 적었습니다.")
            return Filled(target, "새 메뉴 없음", duplicates=len(p.duplicates), overflow=len(p.overflow))
        ui.ok(f"‘{target.sheet}’ {where}에 {len(p.to_write)}개 입력했습니다." + (" 대표 도메인(D6)도 적었습니다." if p.cells else ""))
        return Filled(target, "입력 완료", where, len(p.to_write), len(p.duplicates), len(p.overflow), str(self.out.resolve()))

    def columns(self, sheet):
        return diag.menu_columns(self.current, sheet)

    def clear(self, sheet, items):
        """Empty rows 13~15 of the given columns in the output copy (re-checked against the current file)."""
        now = {c["col"]: c for c in diag.menu_columns(self.current, sheet)}
        ok = [i for i in items if i["col"] in now and now[i["col"]]["menu"] == i["menu"] and not now[i["col"]]["results"]]
        if not ok:
            return {"cleared": [], "skipped": [i["col"] for i in items]}
        p = diag.Plan(sheet, [], 0, cells={f"{diag.L(i['col'])}{r}": "" for i in ok for r in diag.MENU_ROWS})
        if not self._save(p, True):
            return {"cleared": [], "skipped": [i["col"] for i in items]}
        self.current = self.out
        return {"cleared": [i["col"] for i in ok], "skipped": [i["col"] for i in items if i not in ok]}

    def _save(self, p, interactive):
        """Write the plan into the output copy. The copy is replaced only after the new file verified."""
        src = self.src if self.current == self.src else self.out
        tmp = self.out.with_name(self.out.stem + "_저장중" + self.out.suffix)
        try:
            while True:
                try:
                    diag.write(src, tmp, p)
                    os.replace(tmp, self.out)
                    return True
                except PermissionError:
                    msg = f"'{self.out.name}' 파일이 엑셀 등 다른 프로그램에서 열려 있어 저장하지 못했습니다."
                    if interactive and ui.ask_yes(f"  {msg} 파일을 닫고 다시 시도할까요?", True):
                        continue
                    ui.fail(msg + " 이 대상은 기입하지 않았습니다.")
                    return False
                except (OSError, RuntimeError, ValueError) as exc:
                    ui.fail(f"저장하지 못했습니다: {exc}  (이 대상은 기입하지 않았습니다)")
                    return False
        finally:
            if tmp.exists():
                tmp.unlink()
