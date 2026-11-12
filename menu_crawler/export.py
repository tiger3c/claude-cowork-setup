from collections import Counter
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .models import MODES, PHASES

NAVY = "1F3864"
BLUE = "D9E1F2"
YELLOW = "FFF2CC"
ORANGE = "FCE4D6"
GREEN = "E2EFDA"
GRAY = "EDEDED"
LINE = Side(style="thin", color="CBD5E1")


def text_cell(cell, value):
    """All scraped strings stay literal, including =, +, -, @ and XML controls."""
    if isinstance(value, str):
        value = "".join(ch for ch in value if ord(ch) >= 32 or ch in "\t\n\r")[:32767]
        cell.value = value
        cell.data_type = "s"
    else:
        cell.value = value
    return cell


def table(wb, name, headers, rows, widths):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D2"
    for c, header in enumerate(headers, 1):
        cell = text_cell(ws.cell(1, c), header)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    for r, values in enumerate(rows, 2):
        for c, value in enumerate(values, 1):
            cell = text_cell(ws.cell(r, c), value)
            cell.font = Font(name="맑은 고딕", size=10, color="243247")
            cell.fill = PatternFill("solid", fgColor="F3F6FA" if r % 2 == 0 else "FFFFFF")
            cell.border = Border(bottom=LINE)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = 42
    for c, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def select_menus(report, selections=None):
    by_id = {m["id"]: m for m in report["menus"]}
    result = []
    if selections is None:
        selections = [{"id": m["id"], "depths": (m["path"] + ["", "", ""])[:3]} for m in report["menus"] if m["ready"]]
    seen_ids = set()
    for selection in selections:
        ident = selection["id"]
        if ident not in by_id:
            raise ValueError("목록에 없는 메뉴가 선택됐습니다. 화면을 새로고침하세요.")
        if ident in seen_ids:
            continue
        seen_ids.add(ident)
        depths = [" ".join(x.split()) for x in selection["depths"]]
        if len(depths) != 3 or not depths[0] or any(len(x) > 240 for x in depths):
            raise ValueError("선택한 메뉴의 1 Depth를 입력하고 각 단계는 240자 이내로 작성하세요.")
        result.append({**by_id[ident], "depths": depths})
    return result


def make_workbook(report, selections=None):
    selected = select_menus(report, selections)
    wb = Workbook()
    wb.remove(wb.active)
    counts = Counter(m["status"] for m in report["menus"])
    # Header utility and footer links (terms, privacy, sign-up...) stay in 전체메뉴 but are not review work.
    needs = [m for m in report["menus"] if m["evidence"] != "공통 링크"
             and ((not m["ready"] and m["kind"] != "group") or m["hierarchy_review"])]
    mode = report["config"].get("mode", "public")
    compare = mode == "compare"
    visibility = Counter(m.get("visibility") or "" for m in report["menus"])
    summary = [
        ["대상", report["config"]["site_name"]], ["시작 주소", report["config"]["url"]],
        ["수집 방식", MODES.get(mode, mode)],
        ["완료한 단계", ", ".join(PHASES.get(x, x) for x in report.get("phases_done", [])) or "없음"],
        ["수집 상태", {"complete": "설정 범위 탐색 종료", "stopped": "사용자 중지", "ready": "화면 추가·접속 확인 전", "error": "실행 오류"}.get(report["state"], report["state"])],
        ["시작 시각 (UTC)", report["started_at"]], ["종료 시각 (UTC)", report["finished_at"]],
        ["방문 주소 수", report["visited_count"]], ["발견한 메뉴·분류 수", len(report["menus"])],
        ["붙여넣기 선택 메뉴 수", len(selected)], ["미수집·구조 확인 메뉴 수", len(needs)],
        ["페이지 단위 확인 사항", len(report["issues"])], ["수집 제한 도달", "예" if report["limit_reached"] else "아니요"],
        ["최대 메뉴 단계", "3단계 (4단계 이하의 이름·주소 제외)"],
        ["허용 추가 호스트", ", ".join(report["config"]["extra_hosts"]) or "없음 · 시작 호스트만"],
        ["읽는 법", "전체메뉴에서 구조와 상태를 확인하고 미수집_확인필요를 검토하세요. 진단메뉴 시트의 H13부터 15행까지를 기존 진단 시트 H13에 값으로 붙여넣습니다."],
        ["범위 해석", "관찰한 화면·링크를 기준으로 한 목록입니다. 로그인·권한 때문에 화면에 나타나지 않은 메뉴는 이름을 알 수 없어 페이지 단위로 미수집 표시합니다. 탐색 종료가 사이트 전체 메뉴의 완전성을 보장하지는 않습니다."],
        ["수동 검토", "일반 본문 링크는 메뉴 단계 확인 대상으로 분류합니다. 1 Depth만 있는 메뉴도 정상입니다. 수집 상태는 취약점 진단 결과가 아닙니다."],
        ["민감 주소 값", "일반적인 토큰·인증 코드·세션 쿼리 값은 URL에서 [가림]으로 표시합니다."],
    ] + [["상태별 · " + status, count] for status, count in sorted(counts.items())]
    if compare:
        summary += [["노출 · 로그인 전 공개", visibility["로그인 전 공개"]], ["노출 · 로그인 후 전용", visibility["로그인 후 전용"]],
                    ["노출 구분 해석", "‘로그인 후 전용’은 로그인 전 화면에서는 보이지 않고 로그인 후 화면에서만 발견된 메뉴입니다. "
                     "인증·접근통제 점검 대상을 고를 때 참고하세요. 메뉴가 보이지 않는 것이 해당 페이지의 접근 차단을 뜻하지는 않습니다."]]
    ws = table(wb, "수집요약", ["항목", "내용"], summary, [30, 110])
    ws.freeze_panes = "A2"
    for r in range(2, len(summary) + 2):
        if len(str(ws.cell(r, 2).value or "")) > 90:
            ws.row_dimensions[r].height = 62
    inventory = []
    positions = {}
    for i, m in enumerate(report["menus"], 1):
        positions[m["id"]] = i
        depths = (m["path"] + ["", "", ""])[:3]
        inventory.append([i, *depths, " > ".join(m["path"][:3]), m["url"],
                          {"group": "분류 메뉴", "dynamic": "동적 메뉴", "link": "페이지 메뉴"}.get(m["kind"], m["kind"]),
                          m["status"], "필요" if m["hierarchy_review"] else "메뉴 구조 확인", m["evidence"],
                          m["title"], m["http_status"], m["source_url"], m["reason"]]
                         + ([m.get("visibility", ""), m.get("status_before", "")] if compare else []))
    headers = ["No", "1 Depth", "2 Depth", "3 Depth", "메뉴 경로", "URL", "구분", "수집 상태", "구조 검토", "발견 근거", "페이지 제목", "HTTP", "발견 화면", "설명"]
    widths = [7, 24, 28, 32, 58, 65, 16, 20, 20, 22, 40, 9, 65, 65]
    if compare:
        headers += ["노출 구분", "로그인 전 상태"]
        widths += [18, 20]
    ws = table(wb, "전체메뉴", headers, inventory, widths)
    for r, m in enumerate(report["menus"], 2):
        color = GREEN if m["ready"] else GRAY if m["kind"] == "group" and not m["hierarchy_review"] else ORANGE
        for c in (8, 9):
            ws.cell(r, c).fill = PatternFill("solid", fgColor=color)
        if compare and m.get("visibility") == "로그인 후 전용":
            ws.cell(r, 15).fill = PatternFill("solid", fgColor=BLUE)
        ws.row_dimensions[r].height = 44
    issues = [["메뉴", *m["path"][:3], *([""] * (3 - len(m["path"][:3]))), m["url"],
               "구조 확인 필요" if m["status"] == "수집됨" and m["hierarchy_review"] else m["status"],
               m["reason"] or "화면에서 실제 메뉴 단계를 확인하고 선택하세요.", m["source_url"], "", ""] for m in needs]
    issues += [["페이지", "", "", "", x.get("url", ""), x["kind"], x["reason"], "", x.get("phase", ""), x.get("resolved", "")]
               for x in report["issues"]]
    ws = table(wb, "미수집_확인필요", ["구분", "1 Depth", "2 Depth", "3 Depth", "관련 주소", "상태", "확인할 내용", "발견 화면", "수집 단계", "해결 여부"], issues,
               [12, 24, 28, 32, 65, 24, 85, 65, 14, 18])
    for row in ws.iter_rows(min_row=2):
        row[5].fill = PatternFill("solid", fgColor=GREEN if row[9].value else ORANGE)
    if compare:
        only = [m for m in report["menus"] if m.get("visibility") == "로그인 후 전용"]
        table(wb, "로그인후_전용메뉴", ["전체메뉴 No", "1 Depth", "2 Depth", "3 Depth", "URL", "수집 상태", "발견 화면"],
              [[positions[m["id"]], *(m["path"] + ["", "", ""])[:3], m["url"], m["status"], m["source_url"]] for m in only],
              [12, 24, 28, 32, 65, 20, 65])
    table(wb, "방문기록", ["수집 단계", "방문 주소", "최종 주소", "페이지 제목", "HTTP", "상태", "설명"],
          [[p.get(k, "") for k in ["phase", "url", "final_url", "title", "http_status", "status", "reason"]] for p in report["pages"]], [12, 70, 70, 45, 9, 24, 80])
    blocks = [selected[i:i + 250] for i in range(0, len(selected), 250)] or [[]]
    for index, block in enumerate(blocks, 1):
        ws = wb.create_sheet(f"진단메뉴_{index:02}")
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "H17"
        for row, title in [(1, "진단 시트 붙여넣기 · 1·2·3 Depth"), (2, "메뉴명만 복사합니다. URL은 전체메뉴 시트에서 확인하세요."),
                           (3, f"선택 메뉴 {len(selected)}개 · 이 시트 {len(block)}개 · 묶음 {index}/{len(blocks)}"),
                           (5, f"복사 범위: H13:{get_column_letter(7 + len(block))}15 → 기존 진단 시트의 H13에 ‘값만 붙여넣기’" if block else "붙여넣을 메뉴가 없습니다. 전체메뉴·미수집_확인필요를 검토하세요."),
                           (6, "250개가 넘으면 진단메뉴_02부터 추가로 생성합니다. 각 묶음은 별도 진단 시트에 사용하세요. 기존 결과와 연결된 메뉴 열을 덮어쓰지 않도록 확인하세요.")]:
            ws.merge_cells(start_row=row, end_row=row, start_column=1, end_column=16)
            cell = text_cell(ws.cell(row, 1), title)
            cell.font = Font(name="맑은 고딕", size=14 if row == 1 else 10, color="FFFFFF" if row == 1 else NAVY, bold=row in {1, 5})
            cell.fill = PatternFill("solid", fgColor=NAVY if row == 1 else "FFFFFF")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            ws.row_dimensions[row].height = 32
        for r, label in [(12, "메뉴 번호"), (13, "1 Depth · 필수"), (14, "2 Depth · 선택"), (15, "3 Depth · 선택"), (16, "메뉴 경로"), (17, "전체메뉴의 No")]:
            ws.merge_cells(start_row=r, end_row=r, start_column=1, end_column=7)
            text_cell(ws.cell(r, 1), label).font = Font(name="맑은 고딕", size=10, bold=True, color=NAVY)
            ws.cell(r, 1).fill = PatternFill("solid", fgColor=BLUE)
            ws.row_dimensions[r].height = 40 if r >= 13 else 28
        for i, m in enumerate(block, 8):
            values = [(12, (index - 1) * 250 + i - 7), *[(13 + j, value) for j, value in enumerate(m["depths"])],
                      (16, " > ".join(x for x in m["depths"] if x)), (17, positions[m["id"]])]
            for r, value in values:
                cell = text_cell(ws.cell(r, i), value)
                cell.font = Font(name="맑은 고딕", size=10, color=NAVY)
                cell.fill = PatternFill("solid", fgColor=YELLOW if r in {13, 14, 15} else GRAY)
                cell.alignment = Alignment(wrap_text=True, vertical="center")
                cell.border = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
            ws.column_dimensions[get_column_letter(i)].width = 30
        for c in range(1, 8):
            ws.column_dimensions[get_column_letter(c)].width = 8
    output = BytesIO()
    wb.save(output)
    return output.getvalue()
