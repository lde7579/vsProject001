"""나머지 값(WMS 출고, 현재고, Summary 시트)을 읽는다."""
import re
from dataclasses import dataclass

from safety_stock.common.xlsx import (
    InputError,
    find_header_row,
    find_sheet,
    header_map,
    norm_header,
    norm_sku,
    require_cols,
    to_number,
)

_WK_RE = re.compile(r'^WK(\d+)$')

# 표준 키 -> 정규화 제목 (demand는 제목이 시트마다 달라 따로 찾는다)
SUMMARY_FIELDS: dict[str, str] = {
    'group': 'SKU그룹',
    'onhand': '현재고',
    'sigma': '수요표준편차',
    'lt_days': '리드타임(일)',
    'lt_demand': 'LT평균수요',
    'lt_sigma': 'LT수요편차',
    'current_ss_days': '기존안전재고일수',
    'ss_days': '적정안전재고일수',
    'ss': '적정안전재고',
    'target_days': '적정재고일수',
    'rop': 'ROP(재주문점)',
    'correction': '과소예측보정계수',
    'grade': '재고등급',
    'z_coef': 'Z값계수',
}
_TEXT_FIELDS = {'group', 'grade'}


@dataclass(frozen=True)
class WeeklySales:
    """주차별 출고. 주차 열은 왼쪽에서 오른쪽 순서 그대로. 빈 칸은 None."""
    week_numbers: list[int]
    by_sku: dict[str, list[float | None]]


def load_sales(wb) -> WeeklySales:
    """'WMS 출고' 시트(없으면 '출고실적')에서 SKU별 WKnn 열을 읽는다."""
    ws = find_sheet(wb, 'WMS 출고', '출고실적')
    title_row = find_header_row(ws, 'SKU')
    (sku_col,) = require_cols(header_map(ws, title_row), ['SKU'], ws.title)

    week_cols: list[tuple[int, int]] = []  # (열 번호, 주차 숫자)
    for cell in ws[title_row]:
        m = _WK_RE.match(norm_header(cell.value))
        if m:
            week_cols.append((cell.column, int(m.group(1))))
    if not week_cols:
        raise InputError(f"시트 '{ws.title}'에서 주차 열(예: WK01)을 찾을 수 없습니다")

    by_sku: dict[str, list[float | None]] = {}
    for row in ws.iter_rows(min_row=title_row + 1):
        sku = norm_sku(row[sku_col - 1].value) if len(row) >= sku_col else None
        if sku is None or sku in by_sku:
            continue
        by_sku[sku] = [
            to_number(row[c - 1].value) if len(row) >= c else None for c, _ in week_cols
        ]
    return WeeklySales([n for _, n in week_cols], by_sku)


def load_onhand(wb) -> dict[str, float]:
    """이름에 '현재고'가 있고 'Summary'가 없는 첫 시트에서 SKU별 현재고를 합산한다."""
    ws = next(
        (wb[n] for n in wb.sheetnames if '현재고' in n and 'Summary' not in n), None
    )
    if ws is None:
        raise InputError(
            "현재고 시트를 찾을 수 없습니다(이름에 '현재고' 포함, 'Summary' 제외). "
            f"있는 시트: {', '.join(wb.sheetnames)}"
        )
    title_row = find_header_row(ws, '품목 ID')
    sku_col, qty_col = require_cols(
        header_map(ws, title_row), ['품목 ID', '현재고'], ws.title
    )
    result: dict[str, float] = {}
    for row in ws.iter_rows(min_row=title_row + 1):
        sku = norm_sku(row[sku_col - 1].value) if len(row) >= sku_col else None
        if sku is None:
            continue
        qty = to_number(row[qty_col - 1].value) if len(row) >= qty_col else None
        result[sku] = result.get(sku, 0.0) + (qty or 0.0)
    return result


def summary_columns(ws) -> tuple[int, int, dict[str, int | None]]:
    """Summary 시트의 (제목 행, SKU 열, 표준 키 -> 열). 없는 열은 None."""
    title_row = find_header_row(ws, 'SKU')
    hmap = header_map(ws, title_row)
    (sku_col,) = require_cols(hmap, ['SKU'], ws.title)
    cols: dict[str, int | None] = {
        key: hmap.get(norm_header(title)) for key, title in SUMMARY_FIELDS.items()
    }
    # 평균수요 열: '평균수요'를 포함하고 'LT'로 시작하지 않는 첫 열(왼쪽부터)
    cols['demand'] = next(
        (c for k, c in hmap.items() if '평균수요' in k and not k.startswith('LT')), None
    )
    return title_row, sku_col, cols


def load_summary(
    wb, sheet_keyword: str, required: tuple[str, ...] = ()
) -> dict[str, dict[str, float | str | None]]:
    """Summary 시트를 SKU -> {표준 키: 값}으로 읽는다.

    required에 든 표준 키(demand 포함)의 열이 없으면 InputError, 그 외 없는 열의 값은 None.
    """
    ws = find_sheet(wb, sheet_keyword)
    title_row, sku_col, cols = summary_columns(ws)

    missing = [
        SUMMARY_FIELDS.get(k, '평균수요') for k in required if cols.get(k) is None
    ]
    if missing:
        raise InputError(f"시트 '{ws.title}'에 열이 없습니다: {', '.join(missing)}")

    result: dict[str, dict[str, float | str | None]] = {}
    for row in ws.iter_rows(min_row=title_row + 1):
        sku = norm_sku(row[sku_col - 1].value) if len(row) >= sku_col else None
        if sku is None or sku in result:
            continue
        item: dict[str, float | str | None] = {}
        for key, col in cols.items():
            raw = row[col - 1].value if col is not None and len(row) >= col else None
            if key in _TEXT_FIELDS:
                item[key] = None if raw is None else str(raw).strip() or None
            else:
                item[key] = to_number(raw)
        result[sku] = item
    return result
