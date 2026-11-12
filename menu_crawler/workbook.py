"""Write crawled menus into a diagnosis workbook (진단 시트, rows 13-15, columns H:IW).

The workbook is edited at XML level: only the cells that receive a menu name change, everything
else (charts, conditional formats, validations, formulas, other sheets) is copied byte for byte.
openpyxl is not used for saving because it drops charts.

Safety rules
- A column is used only if it is completely empty: no menu name (rows 13-15) and no result (rows 20-67).
- Menus whose path already exists on the sheet are skipped.
- The source file is never overwritten.
"""
import posixpath
import re
from urllib.parse import urlsplit
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree
from openpyxl.utils import column_index_from_string as CI
from openpyxl.utils import get_column_letter as L

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M = "{%s}" % NS
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

FIRST_COL, LAST_COL = CI("H"), CI("IW")          # 250 menu columns
MENU_ROWS = (13, 14, 15)                          # 1 / 2 / 3 Depth
RESULT_ROWS = range(20, 68)                       # 48 evaluation items
SITE_SHEET = re.compile(r"^\d{2,3}_")
REF = re.compile(r"^([A-Za-z]{1,3})(\d+)$")
MARKERS = {"=formula", "=merged"}                 # occupied cells without a usable text


@dataclass
class Plan:
    sheet: str
    free_columns: list[int]
    existing: int
    to_write: list[tuple[int, list[str]]] = field(default_factory=list)
    duplicates: list[list[str]] = field(default_factory=list)
    overflow: list[list[str]] = field(default_factory=list)
    cells: dict = field(default_factory=dict)      # e.g. {"D6": "https://..."}: only cells that are empty


class Workbook:
    """Read-only view of the parts needed to plan a write."""

    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self.names = z.namelist()
            wb = etree.fromstring(z.read("xl/workbook.xml"))
            rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            target = {r.get("Id"): r.get("Target") for r in rels if isinstance(r.tag, str)}
            self.sheets, self.hidden = {}, set()
            for s in wb.iter(M + "sheet"):
                self.sheets[s.get("name")] = _part(target[s.get("{%s}id" % RNS)])
                if s.get("state") in {"hidden", "veryHidden"}:
                    self.hidden.add(s.get("name"))
            self.shared = []
            if "xl/sharedStrings.xml" in self.names:
                sst = etree.fromstring(z.read("xl/sharedStrings.xml"))
                self.shared = [_text_of(si) for si in sst.iter(M + "si")]

    def site_sheets(self):
        return [n for n in self.sheets if SITE_SHEET.match(n) and n.lower() != "00_sample" and n not in self.hidden]

    def root(self, sheet):
        with zipfile.ZipFile(self.path) as z:
            return etree.fromstring(z.read(self.sheets[sheet]))

    def grid(self, sheet, root=None):
        """{(col, row): text} for the menu and result area. Formula cells give their cached text or '=formula';
        cells inside an array/shared formula range or an empty merge are '=merged'/'=formula' (occupied)."""
        root = self.root(sheet) if root is None else root
        out = {}
        area = lambda ci, ri: FIRST_COL <= ci <= LAST_COL and (ri in MENU_ROWS or ri in RESULT_ROWS)  # noqa: E731
        for c, ci, ri in cells(root):
            f = c.find(M + "f")
            if f is not None and f.get("ref") and f.get("t") in {"array", "shared"}:
                for cc, rr in _span(f.get("ref")):
                    if area(cc, rr):
                        out.setdefault((cc, rr), "=formula")
            if not area(ci, ri):
                continue
            value = self._value(c)
            if value not in (None, ""):
                out[(ci, ri)] = value
        for (c1, r1), (c2, r2) in merges(root):
            anchor = out.get((c1, r1), "")
            for cc in range(c1, c2 + 1):
                for rr in range(r1, r2 + 1):
                    if area(cc, rr) and (cc, rr) != (c1, r1):
                        out[(cc, rr)] = anchor if anchor not in ("", None) else "=merged"
            if area(c1, r1) and (c1, r1) not in out and (c2 > c1 or r2 > r1):
                out[(c1, r1)] = "=merged"
        return out

    def _value(self, c):
        """Cached text of a cell; '=formula' for a formula without a cached value."""
        text = self.text(c)
        if text:
            return text
        return "=formula" if c.find(M + "f") is not None else None

    def text(self, c):
        if c.get("t") == "inlineStr":
            node = c.find(M + "is")
            return _text_of(node) if node is not None else ""
        v = c.find(M + "v")
        if v is None or v.text is None:
            return ""
        if c.get("t") == "s":
            try:
                return self.shared[int(v.text)]
            except (ValueError, IndexError):
                return ""
        return v.text


def _part(target):
    """Workbook rels Target -> zip part name ('worksheets/s1.xml', '/xl/worksheets/s1.xml', './worksheets/..')."""
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))


def _text_of(node):
    """Text of <si>/<is>: plain or rich-text runs, without phonetic (rPh) readings."""
    return "".join(t.text or "" for t in node.iter(M + "t") if t.getparent().tag != M + "rPh")


def cells(root):
    """(cell, col, row) for every cell; a missing or lower-case r is derived from the position (allowed by the spec)."""
    sd = root.find(M + "sheetData")
    rnum = 0
    for row in (sd if sd is not None else []):
        if row.tag != M + "row":
            continue
        rnum = int(row.get("r")) if (row.get("r") or "").isdigit() else rnum + 1
        cnum = 0
        for c in row:
            if c.tag != M + "c":
                continue
            m = REF.match(c.get("r") or "")
            cnum = CI(m.group(1).upper()) if m else cnum + 1
            yield c, cnum, (int(m.group(2)) if m else rnum)


def _span(ref):
    a, _, b = ref.partition(":")
    ma, mb = REF.match(a), REF.match(b or a)
    if not ma or not mb:
        return []
    c1, r1, c2, r2 = CI(ma.group(1).upper()), int(ma.group(2)), CI(mb.group(1).upper()), int(mb.group(2))
    return [(c, r) for c in range(c1, min(c2, LAST_COL + 1) + 1) for r in range(r1, min(r2, 80) + 1)]


def merges(root):
    out = []
    for m in root.iter(M + "mergeCell"):
        a, _, b = (m.get("ref") or "").partition(":")
        ma, mb = REF.match(a), REF.match(b or a)
        if ma and mb:
            out.append(((CI(ma.group(1).upper()), int(ma.group(2))), (CI(mb.group(1).upper()), int(mb.group(2)))))
    return out


def hidden_columns(root):
    out = set()
    cols = root.find(M + "cols")
    for c in (cols if cols is not None else []):
        if c.tag == M + "col" and (c.get("hidden") in {"1", "true"} or c.get("width") in {"0", "0.0"}):
            out.update(range(int(c.get("min")), int(c.get("max")) + 1))
    return out


CONTROL = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\ud800-\udfff\ufffe\uffff]")


def clean(text, limit=240):
    """One line, at most `limit` characters, and only characters an xlsx file can hold."""
    return " ".join(CONTROL.sub(" ", str(text or "")).split())[:limit]


def menu_keys(grid):
    """Menu paths already on a sheet (merged 1 Depth labels are copied to the columns they cover)."""
    keys = set()
    for col in {c for c, _ in grid}:
        key = tuple("" if grid.get((col, r), "") in MARKERS else clean(grid.get((col, r), "")) for r in MENU_ROWS)
        if any(key):
            keys.add(key)
    return keys


def plan(path, sheet, menus):
    """menus: list of [d1, d2, d3]. Returns which free column each new menu goes to."""
    wb = Workbook(path)
    if sheet not in wb.sheets:
        raise ValueError(f"시트를 찾을 수 없습니다: {sheet}")
    root = wb.root(sheet)
    grid = wb.grid(sheet, root)
    used = {col for col, _ in grid}
    existing = menu_keys(grid)
    # Like the Google Sheets backends: new menus continue after the last used column, never into gaps
    # (a gap may be left empty on purpose and would break the menu order). Hidden columns are skipped.
    hidden = hidden_columns(root)
    free = [c for c in range(max(used) + 1 if used else FIRST_COL, LAST_COL + 1) if c not in hidden]
    p = Plan(sheet, free, len(existing))
    slots = iter(free)
    seen = set()
    for menu in menus:
        depths = [clean(x) for x in (list(menu) + ["", "", ""])[:3]]
        key = tuple(depths)
        if not depths[0]:
            continue
        if key in existing or key in seen:
            p.duplicates.append(depths)
            continue
        seen.add(key)
        col = next(slots, None)
        if col is None:
            p.overflow.append(depths)
            continue
        p.to_write.append((col, depths))
    return p


def _column_styles(root):
    """{column index: style id} from <cols>, used for cells that do not exist yet."""
    out = {}
    cols = root.find(M + "cols")
    for c in (cols if cols is not None else []):
        if c.get("style"):
            for i in range(int(c.get("min")), int(c.get("max")) + 1):
                out[i] = c.get("style")
    return out


def _set_text(row, ref, text, col_styles=None):
    col = CI(REF.match(ref).group(1))
    cells_ = [c for c in row if c.tag == M + "c"]
    cell = next((c for c in cells_ if c.get("r") == ref), None)
    if cell is None:
        cell = etree.Element(M + "c", r=ref)
        style = row.get("s") if row.get("customFormat") in {"1", "true"} else (col_styles or {}).get(col)
        if style:
            cell.set("s", style)
        after = next((c for c in cells_ if CI(REF.match(c.get("r")).group(1)) > col), None)
        if after is not None:
            after.addprevious(cell)
        elif cells_:
            cells_[-1].addnext(cell)          # before a trailing <extLst>
        else:
            row.insert(0, cell)
        row.attrib.pop("spans", None)
    for child in list(cell):
        cell.remove(child)
    if not text:
        cell.attrib.pop("t", None)
        return
    cell.set("t", "inlineStr")
    is_ = etree.SubElement(cell, M + "is")
    t = etree.SubElement(is_, M + "t")
    t.text = text
    t.set(XML_SPACE, "preserve")


def _normalize(sd):
    """Give every row and cell an explicit upper-case r (same meaning, needed to edit by reference)."""
    rnum = 0
    for row in sd:
        if row.tag != M + "row":
            continue
        rnum = int(row.get("r")) if (row.get("r") or "").isdigit() else rnum + 1
        row.set("r", str(rnum))
        cnum = 0
        for c in row:
            if c.tag != M + "c":
                continue
            m = REF.match(c.get("r") or "")
            cnum = CI(m.group(1).upper()) if m else cnum + 1
            c.set("r", f"{L(cnum)}{rnum}")


def _grow_dimension(root, refs):
    dim = root.find(M + "dimension")
    if dim is None or not refs:
        return
    span = [(CI(m.group(1).upper()), int(m.group(2))) for m in (REF.match(x) for x in (dim.get("ref") or "").split(":")) if m]
    pts = span + [(CI(REF.match(r).group(1)), int(REF.match(r).group(2))) for r in refs]
    c1, r1 = min(c for c, _ in pts), min(r for _, r in pts)
    c2, r2 = max(c for c, _ in pts), max(r for _, r in pts)
    dim.set("ref", f"{L(c1)}{r1}:{L(c2)}{r2}")


def write(src, dst, p: Plan):
    """Write the plan into a copy of src saved as dst. Only the target sheet and workbook.xml change."""
    src, dst = Path(src), Path(dst)
    if src.resolve() == dst.resolve():
        raise ValueError("원본 파일을 덮어쓸 수 없습니다. 다른 파일 이름으로 저장하세요.")
    wb = Workbook(src)
    sheet_part = wb.sheets[p.sheet]
    with zipfile.ZipFile(src) as z:
        root = etree.fromstring(z.read(sheet_part))
        book = etree.fromstring(z.read("xl/workbook.xml"))
    sd = root.find(M + "sheetData")
    _normalize(sd)
    rows = {int(r.get("r")): r for r in sd if r.tag == M + "row"}
    col_styles = _column_styles(root)
    extra_rows = sorted({int(REF.match(ref).group(2)) for ref in p.cells})
    for r in (*MENU_ROWS, *extra_rows):
        if r not in rows:
            new = etree.Element(M + "row", r=str(r))
            after = [x for x in sd if x.tag == M + "row" and int(x.get("r")) < r]
            (after[-1].addnext(new) if after else sd.insert(0, new))
            rows[r] = new
    written = []
    for col, depths in p.to_write:
        for r, text in zip(MENU_ROWS, depths):
            if text:
                _set_text(rows[r], f"{L(col)}{r}", text, col_styles)
                written.append(f"{L(col)}{r}")
    for ref, text in p.cells.items():
        _set_text(rows[int(REF.match(ref).group(2))], ref, clean(text, 2000), col_styles)
        written.append(ref)
    _grow_dimension(root, written)
    # Excel recalculates every formula when the file opens (menu paths and counts depend on these cells).
    calc = book.find(M + "calcPr")
    if calc is None:
        calc = etree.Element(M + "calcPr")
        # CT_Workbook order: sheets, functionGroups, externalReferences, definedNames, calcPr, ...
        anchor = [e for e in book if e.tag in {M + t for t in ("sheets", "functionGroups", "externalReferences", "definedNames")}]
        anchor[-1].addnext(calc)
    calc.set("fullCalcOnLoad", "1")
    new_sheet = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_book = etree.tostring(book, xml_declaration=True, encoding="UTF-8", standalone=True)
    dst.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=dst.parent) as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == sheet_part:
                    zout.writestr(item, new_sheet)
                elif item.filename == "xl/workbook.xml":
                    zout.writestr(item, new_book)
                elif item.filename == "xl/calcChain.xml":
                    continue  # stale after edits; Excel rebuilds it
                else:
                    zout.writestr(item, zin.read(item))
        if "xl/calcChain.xml" in wb.names:
            _drop_calc_chain(tmp_path)
        _verify(tmp_path, p)
        shutil.move(str(tmp_path), dst)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
    return dst


def _verify(path, p):
    """Re-open the written file and check every planned cell; refuse to keep a file that differs."""
    wb = Workbook(path)
    grid = wb.grid(p.sheet)
    for col, depths in p.to_write:
        for r, text in zip(MENU_ROWS, depths):
            if clean(grid.get((col, r), "")) != text:
                raise RuntimeError(f"저장 검증 실패: {L(col)}{r}에 '{text}'가 기록되지 않았습니다. 파일을 만들지 않았습니다.")
    if p.cells:
        got = cell_values(path, p.sheet, p.cells)
        bad = [ref for ref, text in p.cells.items() if got.get(ref) != clean(text)]
        if bad:
            raise RuntimeError(f"저장 검증 실패: {', '.join(bad)}. 파일을 만들지 않았습니다.")


def cell_values(path, sheet, refs):
    """{ref: text} for a few cells of one sheet (cached text for formulas, '=formula' without one)."""
    wb = Workbook(path)
    return {f"{L(ci)}{ri}": clean(wb._value(c) or "", 2000) for c, ci, ri in cells(wb.root(sheet)) if f"{L(ci)}{ri}" in refs}


def next_column(grid, root):
    used = {c for c, _ in grid}
    hidden = hidden_columns(root)
    return next((c for c in range(max(used) + 1 if used else FIRST_COL, LAST_COL + 1) if c not in hidden), LAST_COL + 1)


def _drop_calc_chain(path):
    """Remove calcChain references from [Content_Types].xml and workbook rels after dropping the part."""
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item)
            if item.filename in ("[Content_Types].xml", "xl/_rels/workbook.xml.rels"):
                root = etree.fromstring(data)
                for e in list(root):
                    ref = e.get("PartName") or e.get("Target") or ""
                    if isinstance(e.tag, str) and posixpath.basename(ref).lower() == "calcchain.xml":
                        root.remove(e)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    shutil.move(str(tmp), path)


def menus_from_result(path):
    """Menus from a crawl result file made by this tool (진단메뉴_NN sheets, H13:..15)."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    if not any(n.startswith("진단메뉴_") for n in wb.sheetnames):
        wb.close()
        raise ValueError("이 도구가 만든 수집 결과 엑셀이 아닙니다. (진단메뉴_ 시트가 없습니다)")
    out = []
    for name in sorted(n for n in wb.sheetnames if n.startswith("진단메뉴_")):
        ws = wb[name]
        rows = {r: [c.value for c in row] for r, row in zip(MENU_ROWS, ws.iter_rows(min_row=13, max_row=15, min_col=FIRST_COL, max_col=LAST_COL))}
        for i in range(LAST_COL - FIRST_COL + 1):
            depths = [clean(rows[r][i] if i < len(rows[r]) else "") for r in MENU_ROWS]
            if depths[0]:
                out.append(depths)
    wb.close()
    return out


# Menus whose names come from the site's navigation but whose page could not be confirmed
# (login wall, permission, connection failure, crawl limit). They are real menus for the diagnosis list.
UNVERIFIED = {"로그인 필요", "로그인·권한 확인", "접속 실패", "미방문", "동작 확인 필요"}
SCRIPT_MENU = "동작 확인 필요"            # menu without an address (javascript:/onclick): name is known, page not opened
ACTION_NAME = re.compile(r"(?:로그아웃|logout|log\s*out|sign\s*out|탈퇴|삭제|delete)", re.I)


def menus_from_report(report, include_unverified=False):
    """(confirmed, unverified) menu lists from a crawl report, in tree order."""
    from .export import select_menus
    confirmed = [m["depths"] for m in select_menus(report)]
    unverified = []
    if include_unverified:
        for m in report["menus"]:
            script = m["status"] == SCRIPT_MENU
            if script and (m["kind"] != "dynamic" or ACTION_NAME.search(" ".join(m["path"]))):
                continue                      # logout / delete links are not menus to diagnose
            if (not m["ready"] and m["kind"] in {"link", "dynamic"} and m["evidence"] == "메뉴 DOM" and not m["hierarchy_review"]
                    and 1 <= len(m["path"]) <= 3 and m["status"] in UNVERIFIED):
                unverified.append(([clean(x) for x in (m["path"] + ["", "", ""])[:3]], "주소 없는 메뉴" if script else m["status"]))
        # a parent that only opens its children is not a separate menu column
        paths = [tuple(d) for d in confirmed] + [tuple(d) for d, _ in unverified]
        def parent(d):
            n = len([x for x in d if x])
            return any(len([x for x in o if x]) > n and o[:n] == tuple(d)[:n] for o in paths)
        unverified = [(d, s) for d, s in unverified if not (s == "주소 없는 메뉴" and parent(d))]
    return confirmed, unverified


def unverified_from_result(path):
    """Unverified navigation menus from the 전체메뉴 sheet of a crawl result file."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    out = []
    if "전체메뉴" in wb.sheetnames:
        rows = wb["전체메뉴"].iter_rows(values_only=True)
        head = list(next(rows))
        col = {name: head.index(name) for name in ("1 Depth", "2 Depth", "3 Depth", "구분", "수집 상태", "구조 검토", "발견 근거")}
        for row in rows:
            if (row[col["구분"]] in {"페이지 메뉴", "동적 메뉴"} and row[col["발견 근거"]] == "메뉴 DOM" and row[col["구조 검토"]] != "필요"
                    and row[col["수집 상태"]] in UNVERIFIED and row[col["1 Depth"]]
                    and not ACTION_NAME.search(" ".join(str(row[col[k]] or "") for k in ("1 Depth", "2 Depth", "3 Depth")))):
                out.append(([clean(row[col[k]]) for k in ("1 Depth", "2 Depth", "3 Depth")], row[col["수집 상태"]]))
    wb.close()
    return out


def resolve_sheet(book, query):
    """Exact name, then the number prefix ('06' -> '06_...'), then a unique part of the name."""
    names = book.site_sheets()
    if query in book.sheets:
        return query
    q = query.strip()
    by_number = [n for n in names if n.split("_", 1)[0] == q.zfill(2)] if q.isdigit() else []
    if len(by_number) == 1:
        return by_number[0]
    partial = [n for n in names if q.lower() in n.lower()]
    if len(partial) == 1:
        return partial[0]
    if not partial and not by_number:
        raise ValueError(f"'{query}'에 해당하는 진단 시트가 없습니다.")
    raise ValueError(f"'{query}'에 해당하는 진단 시트가 여러 개입니다: {', '.join((by_number or partial)[:8])}")


def sheet_summary(book, sheet):
    """(existing menu columns, free columns) of a diagnosis sheet."""
    grid = book.grid(sheet)
    used = {c for c, _ in grid}
    menus = {c for c, r in grid if r in MENU_ROWS}
    return len(menus), (LAST_COL - FIRST_COL + 1) - len(used)


def receipt(p, path):
    """Tab-separated record of what was written, skipped as duplicate, or left out for lack of columns."""
    lines = ["구분\t열\t1 Depth\t2 Depth\t3 Depth"]
    lines += [f"입력\t{L(c)}\t" + "\t".join(d) for c, d in p.to_write]
    lines += ["이미 있음\t\t" + "\t".join(d) for d in p.duplicates]
    lines += ["칸 부족\t\t" + "\t".join(d) for d in p.overflow]
    Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    return Path(path)


@dataclass
class PastePlan:
    start: int                      # first column to paste into (H.. IW)
    menus: list[list[str]]
    duplicates: list[list[str]]
    overflow: list[list[str]]
    existing: int
    checked: bool                   # True when the sheet's current state was read from a downloaded copy


def plan_paste(snapshot, sheet, menus, prior=None):
    """Where to paste in a Google Sheets 진단 시트: right after its last used menu column.

    snapshot is the Google sheet downloaded as .xlsx (or None when the sheet is known to be empty).
    Columns to the right of the last used one are all free, so the pasted block is contiguous.
    prior = {"start": next free column, "keys": set of paths} for blocks already pasted in this run.
    """
    existing, start = set(prior["keys"]) if prior else set(), prior["start"] if prior else FIRST_COL
    if snapshot:
        wb = Workbook(snapshot)
        if sheet not in wb.sheets:
            raise ValueError(f"내려받은 파일에 '{sheet}' 시트가 없습니다.")
        root = wb.root(sheet)
        grid = wb.grid(sheet, root)
        if grid:
            start = max(start, next_column(grid, root))
        existing |= menu_keys(grid)
    new, dup, seen = [], [], set()
    for menu in menus:
        depths = [clean(x) for x in (list(menu) + ["", "", ""])[:3]]
        key = tuple(depths)
        if not depths[0]:
            continue
        if key in existing or key in seen:
            dup.append(depths)
            continue
        seen.add(key)
        new.append(depths)
    room = max(0, LAST_COL - start + 1)
    return PastePlan(start, new[:room], dup, new[room:], len(existing), bool(snapshot))


NUMBERISH = re.compile(r"^[\d\s.,:/()%+\-]+$")


def paste_text(menus):
    """3 rows (1/2/3 Depth) x N columns, tab separated; values never start a formula in Sheets."""
    def cell(v):
        v = clean(v)
        # ' keeps the text as typed: no formula, and no number/date conversion of names like 1-2, 0012, 3/4
        return "'" + v if v[:1] in ("=", "+", "-", "@", "'", '"') or NUMBERISH.match(v) else v
    return "\r\n".join("\t".join(cell(m[i]) for m in menus) for i in range(3))


# ---------- 진단 대상 목록 ----------

@dataclass
class Target:
    """One 진단 대상 = one 진단 시트 (e.g. '01_사이트A')."""
    sheet: str
    name: str = ""
    domain: str = ""
    menus: int = 0
    start: int = FIRST_COL        # first column after the last used one
    d6: str = ""                  # 대표 도메인 cell as it is (empty -> can be filled in)
    save_domain: str = ""         # inspection URL to write into an empty D6 together with the menus

    @property
    def addresses(self):
        return domains(self.d6) or ([self.domain] if self.domain else [])

    @property
    def no(self):
        return self.sheet.split("_", 1)[0]

    @property
    def cell(self):
        return f"{L(self.start)}13" if self.start <= LAST_COL else "칸 없음"


HOST = re.compile(r"^(?:[^\W_](?:(?:[^\W_]|-)*[^\W_])?\.)+(?:[^\W\d_]{2,63}|xn--[a-z0-9-]+)$|^localhost$|^\d{1,3}(?:\.\d{1,3}){3}$", re.I)


def valid_host(host):
    """www.example.com, 정부.한국, 10.0.0.1, localhost -- not '2개', 'v1.8', 'a.com(운영'."""
    return bool(host) and bool(HOST.match(host.rstrip(".")))


def address(token):
    """'www.a.com/app' -> 'https://www.a.com/app'; anything that is not a web address -> ''."""
    token = token.strip().strip("'\"<>[]{}.,;:")
    if not token or "@" in token.split("/", 3)[2 if "://" in token else 0]:
        return ""
    url = token if re.match(r"https?://", token, re.I) else ("" if "://" in token else "https://" + token)
    try:
        parts = urlsplit(url)
        host = parts.hostname or ""
        parts.port                                   # noqa: B018 - raises on a bad port
    except ValueError:
        return ""
    return url if valid_host(host) else ""


def domains(value):
    """Every address in a 대표 도메인 cell: one or several, split by spaces, commas, brackets or line breaks.
    Notes such as '(운영)', '입니다.' or the form's '예) ...' hint are ignored."""
    out = []
    text = re.sub(r"예\).*", " ", CONTROL.sub(" ", str(value or "")))
    for token in re.split(r"[\s,;，()（）\[\]]+", text):
        url = address(token)
        if url and url not in out:
            out.append(url)
    return out


def _domain(value):
    found = domains(value)
    return found[0] if found else ""


def targets(path):
    """All 진단 시트 of a workbook with 진단 대상 name (D3), 대표 도메인 (D6) and current menu count."""
    wb = Workbook(path)
    out = []
    for sheet in wb.site_sheets():
        root = wb.root(sheet)
        head = {ri: wb._value(c) or "" for c, ci, ri in cells(root) if ci == 4 and ri in (3, 6)}
        grid = wb.grid(sheet, root)
        menus = {c for (c, r), v in grid.items() if r in MENU_ROWS and v not in MARKERS}
        name = clean(head.get(3, ""))
        if not name or name in MARKERS or name.startswith("예)"):
            name = sheet.split("_", 1)[-1]
        d6 = clean(head.get(6, ""), 2000)
        out.append(Target(sheet, name, _domain(d6), len(menus), next_column(grid, root), "" if d6 in MARKERS else d6))
    return out


def select_targets(text, items):
    """'사이트A, 사이트B' / '전체' (also '01', '01-03') -> targets in the order typed. Numbers are the sheet number prefix."""
    picked, errors = [], []
    for token in [t.strip() for t in re.split(r"[,，]", text or "") if t.strip()]:
        if token in {"전체", "all", "*"}:
            found = list(items)
        elif re.fullmatch(r"\d{1,3}\s*[-~]\s*\d{1,3}", token):
            a, b = sorted(int(x) for x in re.split(r"\s*[-~]\s*", token))
            found = [t for t in items if t.no.isdigit() and a <= int(t.no) <= b]
        elif token.isdigit():
            found = [t for t in items if t.no == token.zfill(2)]
        else:
            q = "".join(token.split()).lower()
            exact = [t for t in items if q in {"".join(t.sheet.split()).lower(), "".join(t.name.split()).lower()}]
            found = exact or [t for t in items if q in "".join((t.sheet + " " + t.name).split()).lower()]
            if len(found) > 1:
                errors.append(f"'{token}'에 해당하는 대상이 여러 개입니다: " + ", ".join(t.sheet for t in found[:6]))
                continue
        if not found:
            errors.append(f"'{token}'에 해당하는 진단 시트가 없습니다.")
        for t in found:
            if t not in picked:
                picked.append(t)
    if errors:
        raise ValueError("\n".join(errors))
    return picked


def _host_path(url):
    try:
        parts = urlsplit(url if "://" in url else "https://" + url)
        port = parts.port
    except ValueError:
        return "", ""
    host = (parts.hostname or "").lower().rstrip(".")
    try:
        host = host.encode("idna").decode("ascii")      # 정부.한국 == xn--or3bn7q.xn--3e0b707e
    except UnicodeError:
        pass
    host = host[4:] if host.startswith("www.") else host
    if port and (parts.scheme.lower(), port) not in {("https", 443), ("http", 80)}:
        host += f":{port}"                                  # a.com:8080 is another service than a.com
    return host, parts.path.rstrip("/").lower()


def match_url(url, items):
    """진단 시트 whose 대표 도메인 (D6) matches the inspection URL: same host, then the longest common path."""
    host, path = _host_path(url)
    if not host:
        return []
    scored = []
    for t in items:
        best = -1
        for address in t.addresses:
            h, p = _host_path(address)
            if h != host:
                continue
            p = p.rsplit("/", 1)[0] if "." in p.rsplit("/", 1)[-1] else p   # /index.jsp -> site root
            if path == p or path.startswith(p + "/") or not p:
                best = max(best, len(p))
        if best >= 0:
            scored.append((best, t))
    if not scored:
        return []
    best = max(n for n, _ in scored)
    return [t for n, t in scored if n == best]


def result_url(path):
    """시작 주소 from the 수집요약 sheet of a crawl result."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True)
    try:
        if "수집요약" in wb.sheetnames:
            for row in wb["수집요약"].iter_rows(values_only=True):
                if row and row[0] == "시작 주소":
                    return str(row[1] or "")
    finally:
        wb.close()
    return ""
