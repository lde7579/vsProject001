"""계산 결과와 엑셀 Summary 값 대조. 입출력 없는 순수 함수."""
from dataclasses import dataclass
import math
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
# 기준과 무관한 입력값(Summary 표준키 -> SkuInputs 속성). 세 기준 행 모두에서 같은 값으로 대조한다
COMPARE_INPUT_FIELDS = {
    'onhand': 'onhand',
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
    return isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)


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
            calc_values = {key: getattr(br, attr) for key, attr in COMPARE_FIELDS.items()}
            calc_values |= {key: getattr(res.inputs, attr) for key, attr in COMPARE_INPUT_FIELDS.items()}
            for key, calc in calc_values.items():
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
    """(기준, SKU)별 판정: '일치' | '불일치' | '비교불가'.

    '불일치' 또는 '계산값없음' 행이 하나라도 있으면 '불일치'(차이를 숨기지 않는다),
    아니고 '일치' 행이 있으면 '일치', 모두 '엑셀값없음'이면 '비교불가'.
    """
    statuses: dict[tuple[str, str], set[str]] = {}
    for r in rows:
        statuses.setdefault((r.basis, r.sku), set()).add(r.status)
    verdicts = {}
    for key, st in statuses.items():
        if st & {'불일치', '계산값없음'}:
            verdicts[key] = '불일치'
        elif '일치' in st:
            verdicts[key] = '일치'
        else:
            verdicts[key] = '비교불가'
    return verdicts
