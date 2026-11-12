"""Odd but valid workbook shapes: the writer must keep the file intact and put menus only where they belong."""
import re
import zipfile

import pytest
from lxml import etree
from openpyxl import load_workbook

from menu_crawler import workbook as diag
from menu_crawler.boards import ExcelBoard
from test_workbook import diagnosis_book

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def patch(src, dst, edits):
    """Copy an xlsx and rewrite parts: edits = {part: function(text) -> text}."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item)
            if item.filename in edits:
                data = edits[item.filename](data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    for part in edits:
        if part not in zipfile.ZipFile(dst).namelist():
            raise AssertionError(part)
    return dst


def sheet_part(path, name):
    return diag.Workbook(path).sheets[name]


def test_output_can_never_be_the_source(tmp_path):
    src = diagnosis_book(tmp_path / "a.xlsx")
    with pytest.raises(ValueError):
        ExcelBoard(src, src)
    with pytest.raises(ValueError):
        ExcelBoard(src, tmp_path / "." / "a.xlsx")
    assert ExcelBoard(src, tmp_path / "b.xlsm").out.suffix == ".xlsx"      # extension follows the content


def test_calcchain_references_removed_in_any_xml_form(tmp_path):
    src = diagnosis_book(tmp_path / "a.xlsx")
    bad = patch(src, tmp_path / "cc.xlsx", {
        "[Content_Types].xml": lambda x: x.replace("</Types>", '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"></Override></Types>'),
        "xl/_rels/workbook.xml.rels": lambda x: x.replace("</Relationships>", '<Relationship Id="rIdCC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"></Relationship></Relationships>'),
    })
    with zipfile.ZipFile(bad, "a") as z:
        z.writestr("xl/calcChain.xml", '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="H16" i="3"/></calcChain>')
    out = diag.write(bad, tmp_path / "o.xlsx", diag.plan(bad, "02_빈시트", [["가"]]))
    with zipfile.ZipFile(out) as z:
        assert "xl/calcChain.xml" not in z.namelist()
        assert b"calcChain" not in z.read("[Content_Types].xml") and b"calcChain" not in z.read("xl/_rels/workbook.xml.rels")


def test_cells_without_r_lowercase_refs_comments_and_extlst(tmp_path):
    src = diagnosis_book(tmp_path / "a.xlsx")
    part = sheet_part(src, "02_빈시트")

    def edit(x):
        x = x.replace("<sheetData>", "<sheetData><!-- generated -->")
        x = re.sub(r'<c r="A13"', '<c', x)                      # cell without r (position = A)
        x = x.replace('r="A14"', 'r="a14"')                     # lower-case reference
        x = re.sub(r'(<row r="15"[^>]*>)(.*?)(</row>)', r'\1\2<extLst/>\3', x)
        return x
    odd = patch(src, tmp_path / "odd.xlsx", {part: edit})
    p = diag.plan(odd, "02_빈시트", [["가", "나", "다"]])
    out = diag.write(odd, tmp_path / "o.xlsx", p)
    wb = load_workbook(out)
    ws = wb["02_빈시트"]
    assert (ws["H13"].value, ws["H14"].value, ws["H15"].value) == ("가", "나", "다")
    assert (ws["A13"].value, ws["A14"].value) == ("1 Depth", "2 Depth")
    row15 = [e for e in etree.fromstring(zipfile.ZipFile(out).read(part)).iter(M + "row") if e.get("r") == "15"][0]
    assert row15[-1].tag == M + "extLst"                       # new cell went before extLst
    assert [t.sheet for t in diag.targets(odd)] == ["01_테스트", "02_빈시트"]


def test_phonetic_text_is_not_part_of_a_menu_name(tmp_path):
    src = diagnosis_book(tmp_path / "a.xlsx")
    part = sheet_part(src, "01_테스트")
    def edit(x):
        return re.sub(r'(<c r="H13"[^>]*>)<is>.*?</is>', r'\1<is><r><t>기</t></r><r><t>존</t></r><rPh sb="0" eb="1"><t>キ</t></rPh></is>', x, count=1)
    odd = patch(src, tmp_path / "ph.xlsx", {part: edit})
    assert "キ" in zipfile.ZipFile(odd).read(part).decode()
    p = diag.plan(odd, "01_테스트", [["기존", "하위"]])
    assert p.to_write == [] and p.duplicates == [["기존", "하위", ""]]


def test_merged_and_hidden_and_array_formula_columns_are_occupied(tmp_path):
    src = diagnosis_book(tmp_path / "a.xlsx")
    wb = load_workbook(src)
    ws = wb["02_빈시트"]
    ws["H13"], ws["H14"], ws["I14"] = "기존", "가입", "탈퇴"
    ws.merge_cells("H13:I13")                 # 1 Depth label over two children
    ws.merge_cells("K13:L13")                 # empty merge
    ws.column_dimensions["N"].hidden = True
    wb.save(src)
    p = diag.plan(src, "02_빈시트", [["기존", "탈퇴"], ["새1"], ["새2"]])
    assert p.duplicates == [["기존", "탈퇴", ""]]
    assert [diag.L(c) for c, _ in p.to_write] == ["M", "O"]   # after the merge, skipping hidden N
    part = sheet_part(src, "01_테스트")
    arr = patch(src, tmp_path / "arr.xlsx", {part: lambda x: x.replace(
        '<c r="H16"', '<c r="P13"><f t="array" ref="P13:R13">1</f><v></v></c><c r="H16"', 1)})
    _ = diag.plan(arr, "01_테스트", [["x"]])


def test_formula_d6_uses_cached_value_and_ports_matter(tmp_path):
    src = diagnosis_book(tmp_path / "a.xlsx")
    part = sheet_part(src, "02_빈시트")
    def edit(x):
        return x.replace("<sheetData>", '<sheetData><row r="6"><c r="D6" t="str"><f>HYPERLINK("https://real.example.com")</f><v>https://real.example.com</v></c></row>', 1)
    odd = patch(src, tmp_path / "f.xlsx", {part: edit})
    items = diag.targets(odd)
    assert [t.sheet for t in diag.match_url("https://real.example.com/x", items)] == ["02_빈시트"]
    a = diag.Target("01_a", d6="https://a.com:8080")
    b = diag.Target("02_b", d6="https://www.a.com")
    assert diag.match_url("https://a.com/", [a, b]) == [b]
    assert diag.match_url("https://a.com:8080/app", [a, b]) == [a]


def test_dimension_grows_and_paste_keeps_numbers_as_text(tmp_path):
    src = diagnosis_book(tmp_path / "a.xlsx")
    out = diag.write(src, tmp_path / "o.xlsx", diag.plan(src, "02_빈시트", [["가"]] * 1 + [["나"]]))
    dim = etree.fromstring(zipfile.ZipFile(out).read(sheet_part(out, "02_빈시트"))).find(M + "dimension").get("ref")
    assert dim.endswith("72") and diag.CI(re.match(r"[A-Z]+", dim.split(":")[1]).group()) >= diag.CI("I")
    text = diag.paste_text([["1-2", "0012", "공지"]])
    assert text.split("\r\n") == ["'1-2", "'0012", "공지"]


def test_three_digit_and_hidden_sheets(tmp_path):
    src = diagnosis_book(tmp_path / "a.xlsx")
    wb = load_workbook(src)
    wb.copy_worksheet(wb["02_빈시트"]).title = "100_세자리"
    wb.copy_worksheet(wb["02_빈시트"]).title = "03_숨김"
    wb["03_숨김"].sheet_state = "hidden"
    wb.save(src)
    names = [t.sheet for t in diag.targets(src)]
    assert "100_세자리" in names and "03_숨김" not in names
