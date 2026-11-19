"""요청자 엑셀의 구매마스터 시트를 읽어 SKU별 행으로 만든다."""
from safety_stock.common.xlsx import (
    InputError, find_header_row, find_sheet, header_cols, header_map, norm_header, norm_sku,
    require_cols, to_number,
)
from safety_stock.purchase_master.model import PurchaseMasterRow, normalize_grade

REQUIRED = ['SKU', '재고등급', '실적 가중치', 'FCST 가중치', '리드타임(일)', '안전재고 값',
            '구매단가', '기초재고', '기말재고', 'FCST', '리드타임', '품질이슈']

# 키 -> (제목, 몇 번째로 나오는 열인지, 0부터)
_COLUMN_SPEC = {
    'sku': ('SKU', 0), 'group': ('품목그룹', 0), 'name': ('품목명', 0),
    'price': ('구매단가', 0), 'begin_stock': ('기초재고', 0), 'end_stock': ('기말재고', 0),
    'fcst_qty': ('FCST', 0), 'lead_time_days': ('리드타임(일)', 0), 'current_ss': ('안전재고 값', 0),
    'grade': ('재고등급', 0), 'weight_actual': ('실적 가중치', 0), 'weight_fcst': ('FCST 가중치', 0),
    'excel_turnover': ('회전율', 0), 'price_grade': ('단가', 0), 'lt_grade': ('리드타임', 0),
    'quality_grade': ('품질이슈', 0), 'total': ('종합', 0),
    'turnover_grade': ('회전율', 1), 'price_score': ('단가', 1), 'lt_score': ('리드타임', 1),
    'quality_score': ('품질이슈', 1),
    'turnover_score': ('회전율', 2),
}


def master_columns(ws, hrow: int) -> dict[str, int | None]:
    """제목과 등장 순서로 구매마스터의 열 번호를 정한다. 없으면 None."""
    found = header_cols(ws, hrow)
    # 등급 열(R~U)과 점수 열(V~Y)이 모두 있어야 첫 번째 등장을 등급, 두 번째를 점수로 읽을 수 있다
    if any(len(found.get(norm_header(t), [])) < 2 for t in ('리드타임', '품질이슈')):
        raise InputError(f"시트 '{ws.title}': 등급 열(단가·회전율·리드타임·품질이슈)과 점수 열이 모두 있어야 합니다")
    result: dict[str, int | None] = {}
    for key, (title, nth) in _COLUMN_SPEC.items():
        cols = found.get(norm_header(title), [])
        result[key] = cols[nth] if len(cols) > nth else None
    return result


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
    require_cols(header_map(ws, hrow), REQUIRED, ws.title)
    cols = master_columns(ws, hrow)

    def cell(row, key):
        col = cols[key]
        return ws.cell(row=row, column=col).value if col else None

    result: dict[str, PurchaseMasterRow] = {}
    for r in range(hrow + 1, ws.max_row + 1):
        sku = norm_sku(cell(r, 'sku'))
        if sku is None or sku in result:
            continue
        result[sku] = PurchaseMasterRow(
            sku=sku,
            group=_text(cell(r, 'group')),
            name=_text(cell(r, 'name')),
            grade=normalize_grade(cell(r, 'grade')),
            weight_actual=to_number(cell(r, 'weight_actual')),
            weight_fcst=to_number(cell(r, 'weight_fcst')),
            lead_time_days=to_number(cell(r, 'lead_time_days')),
            current_ss_days=to_number(cell(r, 'current_ss')),
            price=to_number(cell(r, 'price')),
            begin_stock=to_number(cell(r, 'begin_stock')),
            end_stock=to_number(cell(r, 'end_stock')),
            fcst_qty=to_number(cell(r, 'fcst_qty')),
            lt_grade=normalize_grade(cell(r, 'lt_grade')),
            quality_grade=normalize_grade(cell(r, 'quality_grade')),
            excel_price_grade=normalize_grade(cell(r, 'price_grade')),
            excel_turnover_grade=normalize_grade(cell(r, 'turnover_grade')),
            excel_turnover=to_number(cell(r, 'excel_turnover')),
            excel_total=to_number(cell(r, 'total')),
            sheet_row=r,
        )
    return result
