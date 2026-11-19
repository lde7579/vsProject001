"""8/24 시점 계산값과 이후 실제 출고를 비교하는 사후 검증. 입출력 없는 순수 함수.

실제 출고는 검증에만 쓰고 계산 입력에는 쓰지 않는다.
"""
import math
import statistics
from dataclasses import dataclass

from safety_stock.common.xlsx import InputError
from safety_stock.inventory_days.calc import BASES, SkuResult
from safety_stock.inventory_days.loader import WeeklySales

VALIDATION_BASES = ('기존수기',) + BASES
_NO_GROUP = '미분류'


@dataclass(frozen=True)
class BasisCheck:
    target_qty: float | None
    predicted_week: float | None
    error_week: float | None
    covered: bool | None
    actual_days: float | None


@dataclass(frozen=True)
class ValidationRow:
    sku: str
    group: str | None
    lead_time_days: float | None
    actual_week_avg: float | None
    actual_lt_demand: float | None
    checks: dict[str, BasisCheck]
    memo: str | None


@dataclass(frozen=True)
class SummaryMetric:
    scope: str
    basis: str
    n_demand: int
    wape: float | None
    bias: float | None
    n_lt: int
    coverage: float | None
    median_actual_days: float | None


def lt_window_demand(weekly: list[float], lead_time_days: float | None) -> float | None:
    """리드타임 기간의 실제 수요: 앞 floor(d/7)주 합 + 다음 주 × (d mod 7)/7."""
    if lead_time_days is None or lead_time_days <= 0 or lead_time_days > 7 * len(weekly):
        return None
    full = math.floor(lead_time_days / 7)
    rest = lead_time_days - 7 * full
    total = sum(weekly[:full])
    if rest > 0:
        total += weekly[full] * rest / 7
    return total


def _empty_checks() -> dict[str, BasisCheck]:
    return {b: BasisCheck(None, None, None, None, None) for b in VALIDATION_BASES}


def validate(results: list[SkuResult], actual: WeeklySales, weeks: range) -> list[ValidationRow]:
    """SKU별로 기준마다 목표 재고량·예측 오차·LT 수요 방어 여부·실제 재고일수를 구한다."""
    missing = [w for w in weeks if w not in actual.week_numbers]
    if missing:
        raise InputError(f'실제 출고 파일에 검증 주차가 없습니다: {missing}')
    cols = [actual.week_numbers.index(w) for w in weeks]

    rows = []
    for res in results:
        inp = res.inputs
        series = actual.by_sku.get(inp.sku)
        if series is None:
            rows.append(ValidationRow(inp.sku, inp.group, inp.lead_time_days,
                                      None, None, _empty_checks(), '실적파일없음'))
            continue
        weekly = [(series[i] if i < len(series) and series[i] is not None else 0.0) for i in cols]
        week_avg = statistics.fmean(weekly) if weekly else None
        lt_demand = lt_window_demand(weekly, inp.lead_time_days)

        targets: dict[str, tuple[float | None, float | None]] = {}
        weighted_demand = res.bases['가중평균'].demand_week
        manual = None
        if inp.current_ss_days is not None and weighted_demand is not None:
            manual = inp.current_ss_days * (weighted_demand / 7)
        targets['기존수기'] = (manual, None)
        for b in BASES:
            targets[b] = (res.bases[b].rop, res.bases[b].demand_week)

        checks = {}
        for b, (target, predicted) in targets.items():
            error = predicted - week_avg if predicted is not None and week_avg is not None else None
            covered = target >= lt_demand if target is not None and lt_demand is not None else None
            days = target / (week_avg / 7) if target is not None and week_avg is not None and week_avg > 0 else None
            checks[b] = BasisCheck(target, predicted, error, covered, days)
        rows.append(ValidationRow(inp.sku, inp.group, inp.lead_time_days,
                                  week_avg, lt_demand, checks, None))
    return rows


def _metric(scope: str, basis: str, rows: list[ValidationRow]) -> SummaryMetric:
    pairs = [(r.checks[basis].error_week, r.actual_week_avg) for r in rows
             if r.checks[basis].error_week is not None and r.actual_week_avg is not None]
    total_actual = sum(a for _, a in pairs)
    wape = bias = None
    if pairs and total_actual != 0:
        wape = sum(abs(e) for e, _ in pairs) / total_actual
        bias = sum(e for e, _ in pairs) / total_actual
    covered = [r.checks[basis].covered for r in rows if r.checks[basis].covered is not None]
    days = [r.checks[basis].actual_days for r in rows if r.checks[basis].actual_days is not None]
    return SummaryMetric(
        scope=scope, basis=basis, n_demand=len(pairs), wape=wape, bias=bias,
        n_lt=len(covered), coverage=sum(covered) / len(covered) if covered else None,
        median_actual_days=statistics.median(days) if days else None,
    )


def summarize(rows: list[ValidationRow]) -> list[SummaryMetric]:
    """전체와 품목그룹별로 기준마다 WAPE·편향·방어율·실제 재고일수 중앙값을 낸다."""
    groups: list[str] = []
    for r in rows:
        g = r.group or _NO_GROUP
        if g not in groups:
            groups.append(g)
    out = []
    for scope, subset in [('전체', rows)] + [(g, [r for r in rows if (r.group or _NO_GROUP) == g]) for g in groups]:
        for b in VALIDATION_BASES:
            out.append(_metric(scope, b, subset))
    return out
