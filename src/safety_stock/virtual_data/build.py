"""가상 보완본: 산식과 어긋난 칸만 산식 결과로 채워 새 통합문서를 만든다(발표용 가상 데이터).

원본은 건드리지 않고, 바꾼 칸은 모두 '가상보완_내역' 시트에 남기고, 채우지 못한 칸은 '기록만'으로 구분한다.
"""
import statistics
from dataclasses import dataclass
from pathlib import Path

from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from safety_stock.common.xlsx import InputError, find_header_row, find_sheet, norm_sku, open_workbook, to_number
from safety_stock.inventory_days.base import load_base
from safety_stock.inventory_days.calc import CalcSettings
from safety_stock.inventory_days.compare import compare_with_excel
from safety_stock.inventory_days.loader import summary_columns
from safety_stock.purchase_master.grade import GradeResult, compute_grades, turnover
from safety_stock.purchase_master.loader import load_purchase_master, master_columns
from safety_stock.purchase_master.model import Z_BY_GRADE, PurchaseMasterRow, normalize_grade

LOG_SHEET = '가상보완_내역'
GRADE_REASON = '재고등급 산식 재계산'
# 채우지 않은 행인데 채운 행이 회전율 순위 배열에 들어가면서 달라진 등급 관련 칸의 사유
POPULATION_REASON = '가상 재고로 회전율 순위 모집단 변경'
KIND_CHANGED, KIND_RECORD_ONLY = '변경', '기록만'
# 대조 기준 이름 -> Summary 시트 키워드
SUMMARY_KEYWORDS = {'가중평균': 'Summary_가중평균', '실적기준': 'Summary_실적', 'FCST기준': 'Summary_FCST'}
# 수식 검사 대상: 이 도구가 읽거나 쓰는 시트(이름에 포함된 말)
DATA_SHEET_KEYWORDS = ('구매마스터', 'Summary_', 'WMS 출고', '출고실적', 'SCM FCST', '현재고')


@dataclass(frozen=True)
class Change:
    """바꾼(또는 못 채운) 칸 하나의 기록."""
    sheet: str
    cell: str
    old: object
    new: object
    reason: str
    kind: str  # KIND_CHANGED(칸 값을 바꿈) | KIND_RECORD_ONLY(채우지 못해 사유만 남김, new=None)


def _median_turnover(rows: list[PurchaseMasterRow]) -> float | None:
    values = [t for t in (turnover(r.fcst_qty, r.begin_stock, r.end_stock) for r in rows) if t is not None]
    return statistics.median(values) if values else None


def fill_inventory_gaps(rows: list[PurchaseMasterRow]) -> dict[str, float]:
    """회전율 분모가 0인 SKU의 가상 기초재고(= 기말재고). max(1, round(FCST / 중앙회전율))."""
    median = _median_turnover(rows)
    result = {}
    for r in rows:
        if turnover(r.fcst_qty, r.begin_stock, r.end_stock) is not None:
            continue
        fcst = r.fcst_qty or 0
        result[r.sku] = 1 if fcst <= 0 or not median else max(1, round(fcst / median))
    return result


def _check_source(wb, src: Path, out: Path) -> None:
    if out.resolve() == src.resolve():
        raise InputError(f'결과 파일 경로가 입력 파일과 같습니다: {out}')
    for ws in wb.worksheets:
        if not any(k in ws.title for k in DATA_SHEET_KEYWORDS):
            continue  # 메모 시트 등 읽지 않는 시트의 수식은 그대로 둔다
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('='):
                    raise InputError(f'값만 있는 통합문서가 필요합니다(수식 셀: {ws.title}!{c.coordinate})')


def _set(ws, row: int, col: int, new, reason: str, changes: list[Change]) -> None:
    """칸 값을 바꾸고 내역에 남긴다."""
    cell = ws.cell(row=row, column=col)
    changes.append(Change(ws.title, f'{get_column_letter(col)}{row}', cell.value, new, reason, KIND_CHANGED))
    cell.value = new


def _fill_stock(wb, rows: list[PurchaseMasterRow], changes: list[Change]) -> set[str]:
    """2단계: 회전율 분모 0인 행의 기초·기말재고를 가상 값으로. 채운 SKU 집합을 돌려준다."""
    gaps = fill_inventory_gaps(rows)
    if not gaps:
        return set()
    ws = find_sheet(wb, '구매마스터')
    cols = master_columns(ws, find_header_row(ws, 'SKU'))
    median = _median_turnover(rows)
    note = f'{median:.4f}' if median else '없음'
    reason = f'회전율 분모 0 → 가상 재고(FCST÷중앙회전율 {note})'
    by_sku = {r.sku: r for r in rows}
    for sku, value in gaps.items():
        for key in ('begin_stock', 'end_stock'):
            if cols[key]:
                _set(ws, by_sku[sku].sheet_row, cols[key], value, reason, changes)
    return set(gaps)


def _same(old, new) -> bool:
    if isinstance(new, str):
        return normalize_grade(old) == new
    n = to_number(old)
    return n is not None and abs(n - new) <= 1e-9


def _grade_values(g: GradeResult) -> dict[str, object]:
    return {
        'excel_turnover': None if g.turnover is None else round(g.turnover, 4),
        'price_grade': g.price_grade, 'turnover_grade': g.turnover_grade,
        'price_score': g.scores[0] if g.scores else None, 'turnover_score': g.scores[1] if g.scores else None,
        'lt_score': g.scores[2] if g.scores else None, 'quality_score': g.scores[3] if g.scores else None,
        'total': g.total, 'grade': g.grade,
    }


def _write_grades(wb, rows: list[PurchaseMasterRow], grades: dict[str, GradeResult],
                  before: dict[str, GradeResult], filled: set[str], changes: list[Change]) -> None:
    """3단계: 계산 등급 관련 칸(H, R, S, V~Z, N) 중 계산값과 다른 칸만 쓴다. 등급을 못 구한 행은 건드리지 않는다.

    채우지 않은 행의 칸이 재고를 채우기 전 계산값과 달라졌다면(회전율 순위 배열이 바뀐 탓) 사유를 따로 남긴다.
    """
    ws = find_sheet(wb, '구매마스터')
    cols = master_columns(ws, find_header_row(ws, 'SKU'))
    for r in rows:
        g = grades[r.sku]
        if g.grade is None:
            continue
        prev = _grade_values(before[r.sku])
        for key, new in _grade_values(g).items():
            col = cols[key]
            if col and new is not None and not _same(ws.cell(row=r.sheet_row, column=col).value, new):
                moved = r.sku not in filled and (prev[key] is None or not _same(prev[key], new))
                _set(ws, r.sheet_row, col, new, POPULATION_REASON if moved else GRADE_REASON, changes)


def _sku_rows(ws, title_row: int, sku_col: int) -> dict[str, int]:
    rows: dict[str, int] = {}
    for r in range(title_row + 1, ws.max_row + 1):
        sku = norm_sku(ws.cell(row=r, column=sku_col).value)
        if sku and sku not in rows:
            rows[sku] = r
    return rows


def _fix_summaries(wb, results, compare_rows, grades: dict[str, GradeResult], changes: list[Change]) -> None:
    """4단계: Summary의 불일치·엑셀값없음 칸을 계산값으로, 등급·Z값계수를 계산 등급으로 맞춘다."""
    memo = {r.inputs.sku: ', '.join(r.memo) for r in results}
    layouts = {}
    for basis, keyword in SUMMARY_KEYWORDS.items():
        ws = find_sheet(wb, keyword)
        title_row, sku_col, cols = summary_columns(ws)
        layouts[basis] = (ws, cols, _sku_rows(ws, title_row, sku_col))
    for cr in compare_rows:
        ws, cols, rows = layouts[cr.basis]
        col, row = cols.get(cr.field), rows.get(cr.sku)
        if col is None or row is None:
            continue
        if cr.status in ('불일치', '엑셀값없음') and cr.calc is not None:
            _set(ws, row, col, cr.calc, f'산식 결과로 교체(원래 {cr.status})', changes)
        elif cr.status == '계산값없음':
            changes.append(Change(ws.title, f'{get_column_letter(col)}{row}', ws.cell(row=row, column=col).value,
                                  None, f'산식으로 채울 수 없음({memo.get(cr.sku, "")})',
                                  KIND_RECORD_ONLY))
    for ws, cols, rows in layouts.values():
        for sku, row in rows.items():
            grade = grades[sku].grade if sku in grades else None
            if grade is None:
                continue
            for key, new in (('grade', grade), ('z_coef', Z_BY_GRADE[grade])):
                col = cols.get(key)
                if col and not _same(ws.cell(row=row, column=col).value, new):
                    _set(ws, row, col, new, GRADE_REASON, changes)


def _write_log(wb, changes: list[Change]) -> None:
    if LOG_SHEET in wb.sheetnames:
        del wb[LOG_SHEET]
    ws = wb.create_sheet(LOG_SHEET)
    for c, h in enumerate(['시트', '셀', '구분', '원래 값', '새 값', '사유'], start=1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
    for r, ch in enumerate(changes, start=2):
        for c, v in enumerate([ch.sheet, ch.cell, ch.kind, ch.old, ch.new, ch.reason], start=1):
            ws.cell(row=r, column=c, value=v)
    ws.freeze_panes = 'A2'


def build_virtual(src: Path, out: Path, settings: CalcSettings = CalcSettings()) -> list[Change]:
    """보완본 src를 읽어 산식과 맞는 가상 보완본을 out에 저장하고 바꾼 칸 목록을 돌려준다."""
    src, out = Path(src), Path(out)
    wb = open_workbook(src, data_only=False)
    _check_source(wb, src, out)
    changes: list[Change] = []

    before_rows = list(load_purchase_master(wb).values())
    before = compute_grades(before_rows)  # 재고를 채우기 전 계산 등급(순위 모집단 변경 구분용)
    filled = _fill_stock(wb, before_rows, changes)

    base_data = load_base(wb)  # 재고를 채운 값으로 다시 읽기
    pm, grades = base_data.pm, base_data.grades
    rows = list(pm.values())
    _write_grades(wb, rows, grades, before, filled, changes)

    results = base_data.results(settings)
    _fix_summaries(wb, results, compare_with_excel(results, base_data.summaries), grades, changes)

    _write_log(wb, changes)
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(out)
    except PermissionError:
        raise InputError(f'결과 파일을 저장할 수 없습니다(엑셀에서 열려 있으면 닫고 다시 실행): {out}') from None
    return changes
