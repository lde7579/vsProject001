from pytest import approx

from safety_stock.inventory_days.calc import SkuInputs, compute_sku
from safety_stock.inventory_days.compare import compare_tolerance, compare_with_excel, sku_verdicts


def test_compare_tolerance():
    assert compare_tolerance(134) == approx(0.5 + 0.67, abs=1e-6)
    assert compare_tolerance(1.33) == approx(0.005 + 0.00665, abs=1e-6)
    assert compare_tolerance(0) == approx(0.5, abs=1e-6)


def test_compare_marks_match_mismatch_and_missing():
    res = compute_sku(SkuInputs('X', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 1.5, 50, [10, 20, 30, None], [12.0] * 13))
    w = res.bases['가중평균']
    excel = {'가중평균': {'X': {'demand': round(w.demand_week), 'rop': w.rop + 3, 'target_days': None}},
             '실적기준': {}, 'FCST기준': {}}
    rows = {(r.field): r for r in compare_with_excel([res], excel)}
    assert rows['demand'].status == '일치' and rows['rop'].status == '불일치'
    assert rows['rop'].diff == approx(-3) and rows['target_days'].status == '엑셀값없음'
    assert sku_verdicts(list(rows.values()))[('가중평균', 'X')] == '불일치'


def test_compare_calc_missing_and_sku_not_in_sheet():
    res = compute_sku(SkuInputs('Y', 'G', 'n', 'B', 0.7, 0.3, None, 30, 1.5, 50, [10, 20], [12.0] * 13))
    excel = {'가중평균': {'Y': {'rop': 100}}, '실적기준': {}, 'FCST기준': {}}
    rows = compare_with_excel([res], excel)
    assert {r.basis for r in rows} == {'가중평균'}  # 엑셀 시트에 없는 SKU는 비교하지 않는다
    by = {r.field: r for r in rows}
    assert by['rop'].status == '계산값없음' and by['rop'].diff is None
    assert by['demand'].status == '엑셀값없음'  # 둘 다 없으면 엑셀값없음
    assert sku_verdicts(rows)[('가중평균', 'Y')] == '일치'
