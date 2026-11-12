"""Write crawled menus into a diagnosis workbook (진단 시트, rows 13-15, columns H:IW).

The workbook is edited at XML level: only the cells that receive a menu name change, everything
else (charts, conditional formats, validations, formulas, other sheets) is copied byte for byte.
openpyxl is not used for saving because it drops charts.

Safety rules
- A column is used only if it is completely empty: no menu name (rows 13-15) and no result (rows 20-67).
- Menus whose path already exists on the sheet are skipped.
- The source file is never overwritten.
"""
import re
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
SITE_SHEET = re.compile(r"^\d\d_")
REF = re.compile(r"([A-Z]+)(\d+)$")


@dataclass
class Plan:
    sheet: str
    free_columns: list[int]
    existing: int
    to_write: list[tuple[int, list[str]]] = field(default_factory=list)
    duplicates: list[list[str]] = field(default_factory=list)
    overflow: list[list[str]] = field(default_factory=list)


class Workbook:
    """Read-only view of the parts needed to plan a write."""

    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self.names = z.namelist()
            wb = etree.fromstring(z.read("xl/workbook.xml"))
            rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            target = {r.get("Id"): r.get("Target") for r in rels}
            self.sheets = {}
            for s in wb.iter(M + "sheet"):
                t = target[s.get("{%s}id" % RNS)].lstrip("/")
                self.sheets[s.get("name")] = t if t.startswith("xl/") else "xl/" + t
            self.shared = []
            if "xl/sharedStrings.xml" in self.names:
                sst = etree.fromstring(z.read("xl/sharedStrings.xml"))
                self.shared = ["".join(si.itertext()) for si in sst.iter(M + "si")]

    def site_sheets(self):
        return [n for n in self.sheets if SITE_SHEET.match(n) and n != "00_Sample"]

    def grid(self, sheet):
        """{(col, row): text-or-formula-marker} for the menu and result area."""
        with zipfile.ZipFile(self.path) as z:
            root = etree.fromstring(z.read(self.sheets[sheet]))
        out = {}
        for c in root.iter(M + "c"):
            col, row = REF.match(c.get("r")).groups()
            ci, ri = CI(col), int(row)
            if not (FIRST_COL <= ci <= LAST_COL and (ri in MENU_ROWS or ri in RESULT_ROWS)):
                continue
            value = self._value(c)
            if value not in (None, ""):
                out[(ci, ri)] = value
        return out

    def _value(self, c):
        if c.find(M + "f") is not None:
            return "=formula"
        if c.get("t") == "inlineStr":
            return "".join(c.find(M + "is").itertext())
        v = c.find(M + "v")
        if v is None:
            return None
        if c.get("t") == "s":
            return self.shared[int(v.text)]
        return v.text


def clean(text):
    return " ".join(str(text or "").split())[:240]


def plan(path, sheet, menus):
    """menus: list of [d1, d2, d3]. Returns which free column each new menu goes to."""
    wb = Workbook(path)
    if sheet not in wb.sheets:
        raise ValueError(f"시트를 찾을 수 없습니다: {sheet}")
    grid = wb.grid(sheet)
    used = {col for col, _ in grid}
    existing = set()
    for col in used:
        path_ = tuple(clean(grid.get((col, r), "")) for r in MENU_ROWS)
        if any(path_):
            existing.add(path_)
    free = [c for c in range(FIRST_COL, LAST_COL + 1) if c not in used]
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


def _set_text(row, ref, text):
    col = CI(REF.match(ref).group(1))
    cell = None
    for c in row:
        if c.get("r") == ref:
            cell = c
            break
    if cell is None:
        cell = etree.Element(M + "c", r=ref)
        for i, c in enumerate(row):
            if CI(REF.match(c.get("r")).group(1)) > col:
                row.insert(i, cell)
                break
        else:
            row.append(cell)
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
    rows = {int(r.get("r")): r for r in root.find(M + "sheetData")}
    sd = root.find(M + "sheetData")
    for r in MENU_ROWS:
        if r not in rows:
            new = etree.Element(M + "row", r=str(r))
            after = [x for x in sd if int(x.get("r")) < r]
            (after[-1].addnext(new) if after else sd.insert(0, new))
            rows[r] = new
    for col, depths in p.to_write:
        for r, text in zip(MENU_ROWS, depths):
            if text:
                _set_text(rows[r], f"{L(col)}{r}", text)
    # Excel recalculates every formula when the file opens (menu paths and counts depend on these cells).
    calc = book.find(M + "calcPr")
    if calc is None:
        calc = etree.Element(M + "calcPr")
        anchor = book.find(M + "definedNames")
        if anchor is None:
            anchor = book.find(M + "sheets")
        anchor.addnext(calc)
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
                    zout.writestr(item, zin.read(item.filename))
        if "xl/calcChain.xml" in wb.names:
            _drop_calc_chain(tmp_path)
        shutil.move(str(tmp_path), dst)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
    return dst


def _drop_calc_chain(path):
    """Remove calcChain references from [Content_Types].xml and workbook rels after dropping the part."""
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in ("[Content_Types].xml", "xl/_rels/workbook.xml.rels"):
                data = re.sub(rb"<(Override|Relationship)[^>]*calcChain[^>]*/>", b"", data)
            zout.writestr(item, data)
    shutil.move(str(tmp), path)


def menus_from_result(path):
    """Menus from a crawl result file made by this tool (진단메뉴_NN sheets, H13:..15)."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
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
