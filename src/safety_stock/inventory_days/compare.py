"""계산 결과와 엑셀 Summary 값 대조. 입출력 없는 순수 함수."""
from dataclasses import dataclass
from decimal import Decimal

from safety_stock.inventory_days.calc import BASES, SkuResult

# Summary 표준키 -> BasisResult 속성
COMPARE_FIELDS = {
    'demand': 'demand_week',
    'sigma': 'sigma_week',
    'lt_demand': 'lt_demand',
    'lt_sigma': 'lt_sigma',
    'ss': 'safety_stock',
    'rop': 'rop',
    'target_days': 'target_days',
    'ss_days': 'ss_days',
}


@dataclass(frozen=True)
class CompareRow:
    sku: str
    basis: str
    field: str
    calc: float | None
    excel: float | None
    diff: float | None  # calc - excel
    status: str  # '일치' | '불일치' | '엑셀값없음' | '계산값없음'


def _is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def compare_tolerance(excel_value: float) -> float:
    """엑셀 표시 반올림 범위 + 입력값(보정계수·가중치) 반올림 여유(값의 0.5%)."""
    exponent = Decimal(format(excel_value, '.10g')).as_tuple().exponent
    decimals = max(0, -exponent)
    return 0.5 * 10 ** -decimals + 0.005 * abs(excel_value) + 1e-9


def compare_with_excel(results: list[SkuResult],
                       summaries: dict[str, dict[str, dict]]) -> list[CompareRow]:
    """기준별 엑셀 Summary에 있는 SKU만 필드별로 대조한다."""
    rows = []
    for res in results:
        sku = res.inputs.sku
        for basis in BASES:
            excel_row = summaries.get(basis, {}).get(sku)
            if excel_row is None:
                continue
            br = res.bases[basis]
            for key, attr in COMPARE_FIELDS.items():
                calc = getattr(br, attr)
                excel = excel_row.get(key)
                if not _is_number(excel):
                    rows.append(CompareRow(sku, basis, key, calc, None, None, '엑셀값없음'))
                elif calc is None:
                    rows.append(CompareRow(sku, basis, key, None, excel, None, '계산값없음'))
                else:
                    diff = calc - excel
                    status = '일치' if abs(diff) <= compare_tolerance(excel) else '불일치'
                    rows.append(CompareRow(sku, basis, key, calc, excel, diff, status))
    return rows


def sku_verdicts(rows: list[CompareRow]) -> dict[tuple[str, str], str]:
    """(기준, SKU)별 판정: 필드 중 하나라도 불일치면 '불일치', 아니면 '일치'."""
    verdicts: dict[tuple[str, str], str] = {}
    for r in rows:
        key = (r.basis, r.sku)
        if r.status == '불일치':
            verdicts[key] = '불일치'
        else:
            verdicts.setdefault(key, '일치')
    return verdicts
