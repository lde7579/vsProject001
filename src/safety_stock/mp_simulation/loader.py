"""MP 시뮬레이션 결과(SCM FCST, 기말재고 예상) 시트를 읽는다."""
import re
from dataclasses import dataclass

from safety_stock.common.xlsx import (
    InputError,
    find_header_row,
    find_sheet,
    header_map,
    norm_sku,
    require_cols,
    to_number,
)

_WEEK_RE = re.compile(r'^W\d{1,2}\(\d{2}/\d{2}\)$')


@dataclass(frozen=True)
class WeeklyPlan:
    """주차별 계획. 주차 열은 왼쪽에서 오른쪽 순서 그대로(정렬하지 않음)."""
    sheet: str
    week_labels: list[str]
    by_sku: dict[str, list[float | None]]


def load_weekly_plan(wb, sheet_keyword: str) -> WeeklyPlan:
    """'품목 ID' 제목 행과 다음 행(2단 헤더)에서 주차 열을 찾아 SKU별로 읽는다."""
    ws = find_sheet(wb, sheet_keyword)
    title_row = find_header_row(ws, '품목 ID')

    (sku_col,) = require_cols(header_map(ws, title_row), ['품목 ID'], ws.title)

    # 열 번호 -> 주차 라벨 (제목 행과 다음 행)
    week_cols: dict[int, str] = {}
    label_row = title_row
    for r in (title_row, title_row + 1):
        for cell in ws[r]:
            v = cell.value
            if isinstance(v, str) and _WEEK_RE.match(v.strip()):
                week_cols[cell.column] = v.strip()
                label_row = r
    if not week_cols:
        raise InputError(f"시트 '{ws.title}'에서 주차 열(예: W35(08/24))을 찾을 수 없습니다")
    cols = sorted(week_cols)

    by_sku: dict[str, list[float | None]] = {}
    for row in ws.iter_rows(min_row=label_row + 1):
        sku = norm_sku(row[sku_col - 1].value) if len(row) >= sku_col else None
        if sku is None or sku in by_sku:
            continue
        by_sku[sku] = [
            to_number(row[c - 1].value) if len(row) >= c else None for c in cols
        ]
    return WeeklyPlan(ws.title, [week_cols[c] for c in cols], by_sku)


def first_weeks(plan: WeeklyPlan, n: int) -> dict[str, list[float]]:
    """SKU별 앞 n주 값. 빈 칸은 0으로 채우고, 주차가 n개보다 적으면 있는 만큼."""
    return {
        sku: [0.0 if v is None else v for v in vals[:max(n, 0)]]
        for sku, vals in plan.by_sku.items()
    }
