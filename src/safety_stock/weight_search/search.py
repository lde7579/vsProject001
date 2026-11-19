"""품목그룹별 실적가중치 탐색: 선택 구간 WAPE로 고르고 검증 구간에서 현재와 비교한다.

선택에는 검증 구간 실제 출고를 쓰지 않는다. 입출력은 run_search에서만 한다.
"""
from collections import Counter
from dataclasses import dataclass, replace
from pathlib import Path

from safety_stock.inventory_days.calc import CalcSettings, build_inputs, compute_sku
from safety_stock.inventory_days.loader import load_onhand, load_sales, load_summary
from safety_stock.inventory_days.report import _open
from safety_stock.inventory_days.validate import SummaryMetric, ValidationRow, summarize, validate
from safety_stock.mp_simulation.loader import load_weekly_plan
from safety_stock.purchase_master.grade import compute_grades
from safety_stock.purchase_master.loader import load_purchase_master
from safety_stock.purchase_master.model import PurchaseMasterRow

CANDIDATES = tuple(round(i / 10, 1) for i in range(11))  # 실적가중치 후보, FCST가중치 = 1 - w
MIN_GROUP_SKUS = 5  # 이보다 SKU가 적은 그룹은 공통 최적값을 쓴다
POOLED = '전체'
_NO_GROUP = '미분류'
_TIE = 1e-12

Table = dict[float, dict[str, tuple[int, float | None]]]


def current_weights(pm: dict[str, PurchaseMasterRow]) -> dict[str, float]:
    """그룹별 현재 실적가중치(최빈값, 동수면 작은 값). 그룹이 없는 행과 가중치가 없는 행은 뺀다."""
    found: dict[str, Counter] = {}
    for r in pm.values():
        if r.group is not None and r.weight_actual is not None:
            found.setdefault(r.group, Counter())[r.weight_actual] += 1
    return {g: min(c, key=lambda w: (-c[w], w)) for g, c in found.items()}


def with_weights(pm: dict[str, PurchaseMasterRow], weights: dict[str, float]) -> dict[str, PurchaseMasterRow]:
    """weights에 있는 그룹 중 원래 가중치가 있는 행만 실적·FCST 가중치를 바꾼 사본.

    원래 가중치가 비어 있던 행은 그대로 둬서 현재와 탐색이 같은 SKU 집합으로 계산되게 한다.
    """
    return {sku: replace(r, weight_actual=weights[r.group], weight_fcst=round(1 - weights[r.group], 1))
            if r.group in weights and r.weight_actual is not None and r.weight_fcst is not None else r
            for sku, r in pm.items()}


def group_wape(rows: list[ValidationRow]) -> dict[str, tuple[int, float | None]]:
    """가중평균 수요오차와 실제가 모두 있는 SKU로 '전체'와 그룹별 (SKU 수, WAPE). 실제 합이 0이면 WAPE는 None."""
    acc: dict[str, list[float]] = {}  # 키 -> [SKU 수, Σ|오차|, Σ실제]
    for r in rows:
        err = r.checks['가중평균'].error_week
        if err is None or r.actual_week_avg is None:
            continue
        for key in (POOLED, r.group or _NO_GROUP):
            a = acc.setdefault(key, [0, 0.0, 0.0])
            a[0] += 1
            a[1] += abs(err)
            a[2] += r.actual_week_avg
    return {k: (int(n), err / actual if actual != 0 else None) for k, (n, err, actual) in acc.items()}


@dataclass(frozen=True)
class Choice:
    group: str
    n: int
    current: float
    chosen: float
    method: str  # '그룹최적' | '공통최적' | '현재유지'


def _best(wapes: dict[float, float | None], current: float) -> float | None:
    """WAPE 최소 w. 동점이면 현재 w에 가까운 쪽, 그래도 같으면 작은 w. 값이 하나도 없으면 None."""
    valid = {w: v for w, v in wapes.items() if v is not None}
    if not valid:
        return None
    lowest = min(valid.values())
    return min((w for w, v in valid.items() if v - lowest <= _TIE), key=lambda w: (round(abs(w - current), 9), w))


def choose(table: Table, current: dict[str, float]) -> list[Choice]:
    """w -> group_wape 표에서 그룹별 w를 고른다. 그룹 순서는 current의 순서."""
    def curve(key: str) -> dict[float, float | None]:
        return {w: table[w].get(key, (0, None))[1] for w in table}

    out = []
    for g, cur in current.items():
        n = next((table[w][g][0] for w in table if g in table[w]), 0)
        best = _best(curve(g), cur)
        method = '그룹최적'
        if best is not None and n < MIN_GROUP_SKUS:
            best, method = _best(curve(POOLED), cur), '공통최적'
        if best is None:
            best, method = cur, '현재유지'
        out.append(Choice(g, n, cur, best, method))
    return out


@dataclass(frozen=True)
class SearchResult:
    table: Table
    choices: list[Choice]
    metrics: dict[tuple[str, str], list[SummaryMetric]]  # (구간 '검증'|'전체구간', 가중치 '현재'|'탐색')


def run_search(base: Path, actual: Path, settings: CalcSettings, select_weeks: range,
               holdout_weeks: range, full_weeks: range) -> SearchResult:
    """후보 w마다 모든 그룹에 같은 w를 적용해 선택 구간 WAPE 표를 만들고, 고른 w와 현재 w를 두 구간에서 비교한다."""
    wb, actual_wb = _open(Path(base)), _open(Path(actual))
    pm = load_purchase_master(wb)
    grades = {sku: g.grade for sku, g in compute_grades(list(pm.values())).items()}
    sales, fcst, onhand = load_sales(wb), load_weekly_plan(wb, 'SCM FCST'), load_onhand(wb)
    corr = {sku: row.get('correction')
            for sku, row in load_summary(wb, 'Summary_가중평균', required=('correction',)).items()}
    actual_sales = load_sales(actual_wb)  # 실제 파일은 한 번만 읽는다
    current = current_weights(pm)

    def results(rows):
        return [compute_sku(i) for i in build_inputs(rows, sales, fcst, onhand, corr, settings, grades=grades)]

    table: Table = {w: group_wape(validate(results(with_weights(pm, dict.fromkeys(current, w))),
                                           actual_sales, select_weeks)) for w in CANDIDATES}
    choices = choose(table, current)
    applied = {c.group: c.chosen for c in choices if c.method != '현재유지'}  # 현재유지 그룹은 원래 행 그대로
    variants = {'현재': results(pm), '탐색': results(with_weights(pm, applied))}
    metrics = {(label, name): summarize(validate(res, actual_sales, weeks))
               for label, weeks in (('검증', holdout_weeks), ('전체구간', full_weeks))
               for name, res in variants.items()}
    return SearchResult(table, choices, metrics)
