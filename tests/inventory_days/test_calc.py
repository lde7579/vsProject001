import math
import statistics

from pytest import approx

from safety_stock.inventory_days.validate import validate

from safety_stock.inventory_days.calc import (
    CalcSettings,
    SkuInputs,
    build_inputs,
    compute_basis,
    compute_sku,
)
from safety_stock.inventory_days.loader import WeeklySales
from safety_stock.mp_simulation.loader import WeeklyPlan
from safety_stock.purchase_master.model import PurchaseMasterRow


def test_compute_basis_matches_excel_reference():
    # 설계 7장 기준값: 요청자 엑셀 ACL131T0CJDG (수요 191, σ 59, LT 100일, C등급, 현재고 722)
    b = compute_basis(191, 59, 100, 'C', 722)
    assert b.lt_demand == approx(191 * 100 / 7)
    assert b.lt_sigma == approx(59 * math.sqrt(100 / 7))
    assert b.safety_stock == approx(b.lt_sigma * 1.036)
    assert b.rop == approx(b.lt_demand + b.safety_stock)
    assert round(b.target_days) == 108 and round(b.ss_days) == 8 and b.order == 'ORDER'


def test_compute_sku_three_bases():
    inp = SkuInputs(sku='X', group='G', name='n', grade='B', weight_actual=0.7, weight_fcst=0.3,
                    lead_time_days=14, current_ss_days=30, correction=1.5, onhand=50,
                    actual_hist=[10, 20, 30, None], fcst13=[12.0] * 13)
    r = compute_sku(inp)
    sd = statistics.pstdev([10, 20, 30])
    assert r.actual_mean == 20 and r.fcst_mean == 12
    w = r.bases['가중평균']
    assert w.demand_week == approx(0.7 * 20 + 0.3 * 12) and w.sigma_week == approx(0.7 * sd)
    assert w.target_days == approx(w.rop / (w.demand_week / 7)) and w.order == 'OK'
    assert r.bases['실적기준'].demand_week == approx(30) and r.bases['실적기준'].sigma_week == approx(sd)
    assert r.bases['FCST기준'].demand_week == approx(18) and r.bases['FCST기준'].sigma_week == 0
    assert r.memo == []


def test_compute_sku_memos_and_blanks():
    inp = SkuInputs(sku='Y', group=None, name=None, grade=None, weight_actual=0.6, weight_fcst=0.3,
                    lead_time_days=None, current_ss_days=None, correction=None, onhand=None,
                    actual_hist=None, fcst13=None)
    r = compute_sku(inp)
    assert set(r.memo) == {'등급없음', '수요없음', '실적없음', 'FCST없음', '가중치합오류', '리드타임없음', '보정계수없음'}
    assert r.bases['가중평균'].demand_week == 0 and r.bases['가중평균'].target_days is None
    assert r.bases['실적기준'].demand_week is None


def test_build_inputs_selects_weeks_and_union():
    # 출고 WK30~WK32 중 WK31까지만, 구매마스터에 없는 출고 SKU도 포함
    sales = WeeklySales(week_numbers=[30, 31, 32], by_sku={'A': [1, 2, 3], 'Z': [5, 5, 5]})
    pm = {'A': PurchaseMasterRow('A', 'G', 'n', 'A', 0.7, 0.3, 14, 30)}
    fcst = WeeklyPlan('f', ['W35(08/24)'], {})
    inputs = build_inputs(pm, sales, fcst, {}, {}, CalcSettings())
    assert [i.sku for i in inputs] == ['A', 'Z'] and inputs[0].actual_hist == [1, 2]
    assert inputs[0].fcst13 is None
    assert 'FCST없음' in compute_sku(inputs[0]).memo and '구매마스터없음' in compute_sku(inputs[1]).memo


def test_build_inputs_fcst_only_sku_and_order():
    sales = WeeklySales(week_numbers=[1], by_sku={'S': [4]})
    pm = {'M': PurchaseMasterRow('M', None, None, 'C', 0.5, 0.5, 7, None)}
    fcst = WeeklyPlan('f', ['W35(08/24)', 'W36(08/31)'], {'F': [1, None], 'M': [2, 3]})
    inputs = build_inputs(pm, sales, fcst, {'M': 9.0}, {'M': 1.2}, CalcSettings(fcst_weeks=2))
    assert [i.sku for i in inputs] == ['M', 'S', 'F']
    assert inputs[0].fcst13 == [2, 3] and inputs[0].onhand == 9.0 and inputs[0].correction == 1.2
    assert inputs[2].fcst13 == [1, 0] and inputs[2].in_master is False


def test_z_none_when_demand_none():
    inp = SkuInputs(sku='X', group=None, name=None, grade='B', weight_actual=0.7, weight_fcst=0.3,
                    lead_time_days=14, current_ss_days=None, correction=None, onhand=1,
                    actual_hist=[1, 2], fcst13=[1.0] * 13)
    r = compute_sku(inp)
    assert r.bases['실적기준'].z is None and r.bases['FCST기준'].z is None
    assert r.bases['가중평균'].z == 1.282


def test_negative_lead_time_does_not_raise():
    inp = SkuInputs(sku='X', group=None, name=None, grade='B', weight_actual=0.7, weight_fcst=0.3,
                    lead_time_days=-7, current_ss_days=None, correction=1.0, onhand=1,
                    actual_hist=[1, 2], fcst13=[1.0] * 13)
    r = compute_sku(inp)
    assert r.bases['가중평균'].lt_weeks is None and r.bases['가중평균'].rop is None
    assert '리드타임없음' in r.memo


def test_fcst_none_cells_count_as_zero():
    inp = SkuInputs(sku='X', group=None, name=None, grade='B', weight_actual=0.0, weight_fcst=1.0,
                    lead_time_days=7, current_ss_days=None, correction=1.0, onhand=1,
                    actual_hist=[1], fcst13=[10.0, None])
    assert compute_sku(inp).fcst_mean == 5.0


def _inp(**kw):
    base = dict(sku='X', group='G', name='n', grade='B', weight_actual=0.7, weight_fcst=0.3,
                lead_time_days=14, current_ss_days=30, correction=1.0, onhand=50,
                actual_hist=[10.0, 20.0], fcst13=[12.0] * 13)
    base.update(kw)
    return SkuInputs(**base)


def test_weight_none_memo_differs_from_sum_error():
    assert '가중치없음' in compute_sku(_inp(weight_fcst=None)).memo
    assert '가중치합오류' not in compute_sku(_inp(weight_fcst=None)).memo
    assert '가중치합오류' in compute_sku(_inp(weight_actual=0.6)).memo
    assert '가중치없음' not in compute_sku(_inp(weight_actual=0.6)).memo


def test_abnormal_correction_memo_keeps_calculation():
    r = compute_sku(_inp(correction=29.82))
    assert '보정계수이상' in r.memo
    assert r.bases['실적기준'].demand_week == approx(15 * 29.82)  # 계산은 그대로
    assert '보정계수이상' in compute_sku(_inp(correction=0.0)).memo
    assert '보정계수이상' not in compute_sku(_inp(correction=3.0)).memo
    assert '보정계수이상' not in compute_sku(_inp(correction=None)).memo


def test_all_blank_actual_rows_count_as_no_actual():
    # 출고 시트에 SKU 행은 있지만 실적 주차가 전부 빈 칸이면 실적없음
    r = compute_sku(_inp(actual_hist=[None, None]))
    assert '실적없음' in r.memo and r.actual_mean is None
    assert r.bases['가중평균'].demand_week == approx(0.3 * 12)  # 계산은 기존대로(실적 0)
    # Task 7 규칙: 실적없음이면 기존수기·가중평균·실적기준은 채점 제외, FCST기준은 채점
    [row] = validate([r], WeeklySales([35], {'X': [10.0]}), range(35, 36))
    for b in ('기존수기', '가중평균', '실적기준'):
        assert row.checks[b].excluded == '실적없음' and row.checks[b].error_week is None
    assert row.checks['FCST기준'].excluded is None and row.checks['FCST기준'].error_week is not None
