"""세 기준(가중평균·실적기준·FCST기준) 적정 재고일수 계산. 입출력 없는 순수 함수."""
import math
import statistics
from dataclasses import dataclass

from safety_stock.mp_simulation.loader import WeeklyPlan, first_weeks
from safety_stock.inventory_days.loader import WeeklySales
from safety_stock.purchase_master.model import PurchaseMasterRow, z_for_grade

BASES = ('가중평균', '실적기준', 'FCST기준')


@dataclass(frozen=True)
class CalcSettings:
    actual_first_week: int = 1
    actual_last_week: int = 31
    fcst_weeks: int = 13


@dataclass(frozen=True)
class SkuInputs:
    sku: str
    group: str | None
    name: str | None
    grade: str | None
    weight_actual: float | None
    weight_fcst: float | None
    lead_time_days: float | None
    current_ss_days: float | None
    correction: float | None
    onhand: float | None
    actual_hist: list[float | None] | None  # None: 출고 시트에 SKU 없음
    fcst13: list[float] | None  # None: FCST 시트에 SKU 없음
    in_master: bool = True


@dataclass(frozen=True)
class BasisResult:
    demand_week: float | None
    sigma_week: float | None
    lt_weeks: float | None
    lt_demand: float | None
    lt_sigma: float | None
    z: float | None
    safety_stock: float | None
    rop: float | None
    target_days: float | None
    ss_days: float | None
    onhand_days: float | None
    order: str | None


@dataclass(frozen=True)
class SkuResult:
    inputs: SkuInputs
    actual_mean: float | None
    fcst_mean: float | None
    bases: dict[str, BasisResult]
    memo: list[str]


def compute_basis(demand_week, sigma_week, lead_time_days, grade, onhand) -> BasisResult:
    """주간 수요·표준편차로 LT 수요, 안전재고, ROP, 일수를 계산한다(반올림 없음)."""
    if demand_week is None:
        return BasisResult(None, None, None, None, None, None, None, None, None, None, None, None)
    z = z_for_grade(grade)

    lt_weeks = lt_demand = lt_sigma = None
    if lead_time_days is not None and lead_time_days > 0:  # 없음·0·음수면 LT 관련 값은 모두 None
        lt_weeks = lead_time_days / 7
        lt_demand = demand_week * lt_weeks
        if sigma_week is not None:
            lt_sigma = sigma_week * math.sqrt(lt_weeks)

    safety_stock = rop = None
    if z is not None and lt_demand is not None and lt_sigma is not None:
        safety_stock = lt_sigma * z
        rop = lt_demand + safety_stock

    target_days = ss_days = onhand_days = None
    if demand_week != 0:
        daily = demand_week / 7
        if rop is not None:
            target_days = rop / daily
            ss_days = safety_stock / daily
        if onhand is not None:
            onhand_days = onhand / daily

    order = None
    if rop is not None and onhand is not None:
        order = 'ORDER' if onhand <= rop else 'OK'

    return BasisResult(demand_week, sigma_week, lt_weeks, lt_demand, lt_sigma, z,
                       safety_stock, rop, target_days, ss_days, onhand_days, order)


def _mean_std(values) -> tuple[float, float]:
    """None을 뺀 평균·모표준편차(엑셀 AVERAGE/STDEV.P). 값이 없으면 0, 0."""
    vals = [v for v in (values or []) if v is not None]
    if not vals:
        return 0.0, 0.0
    return statistics.fmean(vals), statistics.pstdev(vals)


# 보정계수가 이 값보다 크거나 0이면 점검메모 '보정계수이상'(계산은 그대로)
CORRECTION_MAX = 3.0


def compute_sku(inp: SkuInputs) -> SkuResult:
    """SKU 하나의 세 기준 결과와 점검메모."""
    hm, hs = _mean_std(inp.actual_hist)
    # 출고 시트에 행이 없거나, 행은 있어도 실적 주차 값이 모두 빈 칸이면 실적이 없는 것으로 본다
    has_actual = inp.actual_hist is not None and any(v is not None for v in inp.actual_hist)
    # FCST는 빈 칸을 0으로 보고 주 수 전체로 평균·표준편차를 구한다
    fm, fs = _mean_std([0.0 if v is None else v for v in inp.fcst13] if inp.fcst13 else None)
    wh, wf, k = inp.weight_actual, inp.weight_fcst, inp.correction

    if wh is not None and wf is not None:
        weighted = (hm * wh + fm * wf, hs * wh + fs * wf)
    else:
        weighted = (None, None)
    actual = (hm * k, hs) if k is not None else (None, None)
    fcst = (fm * k, fs) if k is not None else (None, None)

    args = (inp.lead_time_days, inp.grade, inp.onhand)
    bases = {
        '가중평균': compute_basis(*weighted, *args),
        '실적기준': compute_basis(*actual, *args),
        'FCST기준': compute_basis(*fcst, *args),
    }

    memo: list[str] = []
    if not inp.in_master:
        memo.append('구매마스터없음')
    if z_for_grade(inp.grade) is None:
        memo.append('등급없음')
    if weighted[0] == 0:
        memo.append('수요없음')
    if not has_actual:
        memo.append('실적없음')
    if inp.fcst13 is None:
        memo.append('FCST없음')
    if wh is None or wf is None:
        memo.append('가중치없음')
    elif abs(wh + wf - 1) > 1e-9:
        memo.append('가중치합오류')
    if inp.lead_time_days is None or inp.lead_time_days <= 0:
        memo.append('리드타임없음')
    if k is None:
        memo.append('보정계수없음')
    elif k == 0 or k > CORRECTION_MAX:
        memo.append('보정계수이상')

    return SkuResult(
        inputs=inp,
        actual_mean=hm if has_actual else None,
        fcst_mean=fm if inp.fcst13 is not None else None,
        bases=bases,
        memo=memo,
    )


def build_inputs(pm: dict[str, PurchaseMasterRow], sales: WeeklySales, fcst: WeeklyPlan,
                 onhand: dict[str, float], corrections: dict[str, float | None],
                 settings: CalcSettings, *, grades: dict[str, str | None]) -> list[SkuInputs]:
    """구매마스터 → 출고에만 있는 SKU → FCST에만 있는 SKU 순서로 입력을 만든다.

    재고등급은 엑셀 값이 아니라 넘겨받은 계산 등급(grades)을 쓴다. 없으면 None(등급없음).
    """
    cols = [i for i, w in enumerate(sales.week_numbers)
            if settings.actual_first_week <= w <= settings.actual_last_week]
    fcst_first = first_weeks(fcst, settings.fcst_weeks)

    skus = list(pm)
    seen = set(skus)
    for sku in list(sales.by_sku) + list(fcst.by_sku):
        if sku not in seen:
            seen.add(sku)
            skus.append(sku)

    result = []
    for sku in skus:
        row = pm.get(sku)
        hist = sales.by_sku.get(sku)
        result.append(SkuInputs(
            sku=sku,
            group=row.group if row else None,
            name=row.name if row else None,
            grade=grades.get(sku),
            weight_actual=row.weight_actual if row else None,
            weight_fcst=row.weight_fcst if row else None,
            lead_time_days=row.lead_time_days if row else None,
            current_ss_days=row.current_ss_days if row else None,
            correction=corrections.get(sku),
            onhand=onhand.get(sku),
            actual_hist=[hist[i] for i in cols if i < len(hist)] if hist is not None else None,
            fcst13=fcst_first.get(sku),
            in_master=row is not None,
        ))
    return result
