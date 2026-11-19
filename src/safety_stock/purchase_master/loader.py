"""요청자 엑셀의 구매마스터 시트를 읽어 SKU별 행으로 만든다."""
from safety_stock.common.xlsx import (
    find_header_row, find_sheet, header_map, norm_header, norm_sku,
    require_cols, to_number,
)
from safety_stock.purchase_master.model import PurchaseMasterRow, normalize_grade

REQUIRED = ['SKU', '재고등급', '실적 가중치', 'FCST 가중치', '리드타임(일)', '안전재고 값']


def _text(v) -> str | None:
    """셀 값을 공백 제거한 문자열로. 비면 None."""
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def load_purchase_master(wb) -> dict[str, PurchaseMasterRow]:
    """구매마스터 시트를 읽는다. 같은 SKU가 또 나오면 첫 행을 유지한다."""
    ws = find_sheet(wb, '구매마스터')
    hrow = find_header_row(ws, 'SKU')
    hmap = header_map(ws, hrow)
    c_sku, c_grade, c_wa, c_wf, c_lt, c_ss = require_cols(hmap, REQUIRED, ws.title)
    c_group = hmap.get(norm_header('품목그룹'))
    c_name = hmap.get(norm_header('품목명'))

    def cell(row, col):
        return ws.cell(row=row, column=col).value if col else None

    result: dict[str, PurchaseMasterRow] = {}
    for r in range(hrow + 1, ws.max_row + 1):
        sku = norm_sku(cell(r, c_sku))
        if sku is None or sku in result:
            continue
        result[sku] = PurchaseMasterRow(
            sku=sku,
            group=_text(cell(r, c_group)),
            name=_text(cell(r, c_name)),
            grade=normalize_grade(cell(r, c_grade)),
            weight_actual=to_number(cell(r, c_wa)),
            weight_fcst=to_number(cell(r, c_wf)),
            lead_time_days=to_number(cell(r, c_lt)),
            current_ss_days=to_number(cell(r, c_ss)),
        )
    return result
