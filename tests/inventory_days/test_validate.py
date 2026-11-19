import pytest
from pytest import approx

from safety_stock.common.xlsx import InputError
from safety_stock.inventory_days.calc import SkuInputs, compute_sku
from safety_stock.inventory_days.loader import WeeklySales
from safety_stock.inventory_days.validate import lt_window_demand, summarize, validate


def test_lt_window_demand():
    w = [7.0] * 9
    assert lt_window_demand(w, 14) == 14 and lt_window_demand(w, 17) == approx(17)
    assert lt_window_demand(w, 63) == 63 and lt_window_demand(w, 64) is None and lt_window_demand(w, None) is None
    assert lt_window_demand(w, 0) is None


def test_validate_and_summarize():
    res = compute_sku(SkuInputs('X', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 1.5, 50, [10, 20, 30, None], [12.0] * 13))
    actual = WeeklySales(week_numbers=list(range(35, 44)), by_sku={'X': [20.0] * 9})
    [row] = validate([res, ], actual, range(35, 44))
    assert row.actual_week_avg == 20 and row.actual_lt_demand == 40
    w = row.checks['가중평균']
    assert w.error_week == approx(17.6 - 20) and w.covered is True
    assert row.checks['기존수기'].target_qty == approx(17.6 * 2 + 30 * 17.6 / 7) and row.checks['기존수기'].predicted_week is None
    m = {(s.scope, s.basis): s for s in summarize([row])}
    assert m[('전체', '가중평균')].wape == approx(2.4 / 20) and m[('전체', '가중평균')].coverage == 1.0
    assert m[('G', '기존수기')].n_demand == 0 and m[('G', '기존수기')].wape is None


def test_validate_missing_sku_in_actual():
    res = compute_sku(SkuInputs('BIDS51DR15WH', 'BID', 'n', 'B', 0.7, 0.3, 42, 45, 1.0, 10, [5.0], [5.0] * 13))
    row, only = validate([res], WeeklySales([35], {'BIDSS1DR15WH': [1.0]}), range(35, 36))
    assert row.memo == '실적파일없음' and row.actual_week_avg is None
    # 실제 파일 쪽 SKU도 사유와 함께 남는다(설계 8장)
    assert only.sku == 'BIDSS1DR15WH' and only.memo == '계산대상없음' and only.actual_week_avg == 1.0


def test_validate_missing_weeks_raises():
    res = compute_sku(SkuInputs('X', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 1.0, 50, [10.0], [12.0] * 13))
    with pytest.raises(InputError, match='43'):
        validate([res], WeeklySales([35, 36], {'X': [1.0, 2.0]}), range(35, 44))


def test_validate_none_as_zero_and_uses_week_columns():
    res = compute_sku(SkuInputs('X', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 1.0, 50, [10.0], [10.0] * 13))
    actual = WeeklySales([34, 35, 36], {'X': [99.0, None, 14.0]})
    [row] = validate([res], actual, range(35, 37))
    assert row.actual_week_avg == 7 and row.actual_lt_demand == 14
    assert row.checks['실적기준'].actual_days == approx(row.checks['실적기준'].target_qty / 1.0)


def test_validate_empty_weeks_raises():
    with pytest.raises(InputError):
        validate([], WeeklySales([35], {}), range(35, 35))


def _three_skus():
    ok = compute_sku(SkuInputs('A', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 1.5, 50, [10, 20, 30, None], [12.0] * 13))
    nograde = compute_sku(SkuInputs('N', 'G', 'n', None, 0.7, 0.3, 14, 30, 1.0, 50, [10.0, 10.0], [10.0] * 13))
    nofcst = compute_sku(SkuInputs('F', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 1.0, 50, [10.0, 10.0], None))
    return ok, nograde, nofcst


def test_excluded_when_input_missing():
    _, _, nofcst = _three_skus()
    actual = WeeklySales([35, 36], {'F': [10.0, 10.0]})
    [row] = validate([nofcst], actual, range(35, 37))
    for b in ('FCST기준', '가중평균', '기존수기'):
        c = row.checks[b]
        assert c.excluded == 'FCST없음' and c.covered is None and c.predicted_week is None and c.target_qty is None
    assert row.checks['실적기준'].excluded is None and row.checks['실적기준'].covered is not None


def test_common_scope_and_aggregation():
    ok, nograde, nofcst = _three_skus()
    b = compute_sku(SkuInputs('B', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 1.0, 50, [10.0, 10.0], [10.0] * 13))
    actual = WeeklySales([35, 36], {'A': [20.0, 20.0], 'N': [10.0, 10.0], 'F': [10.0, 10.0], 'B': [5.0, 5.0]})
    rows = validate([ok, nograde, nofcst, b], actual, range(35, 37))
    m = {(s.scope, s.basis): s for s in summarize(rows)}
    # 정상 A(오차 -2.4, 실제 20)와 B(오차 +5, 실제 5): 합 기준 비율이지 비율 평균이 아니다
    w = m[('공통', '가중평균')]
    assert w.n_demand == 3  # A, B, 등급없음 N (FCST 없는 F는 제외)
    assert w.wape == approx((2.4 + 5 + 0) / 35) and w.bias == approx((-2.4 + 5 + 0) / 35)
    assert m[('공통', '기존수기')].n_demand == 0 and m[('공통', '기존수기')].wape is None
    assert all(m[('공통', x)].n_lt == 2 for x in ('기존수기', '가중평균', '실적기준', 'FCST기준'))
    assert m[('전체', '기존수기')].n_lt == 3  # N도 기존수기는 채점됨
    order = [s.scope for s in summarize(rows)]
    assert order[0] == '전체' and order[4] == '공통' and order[8] == 'G'


def test_group_scope_uses_same_skus_across_bases():
    # 그룹 G 안에 FCST가 없는 F(FCST기준·가중평균·기존수기 채점 불가)가 있어도 기준마다 SKU 수가 같다
    ok, nograde, nofcst = _three_skus()
    b = compute_sku(SkuInputs('B', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 1.0, 50, [10.0, 10.0], [10.0] * 13))
    actual = WeeklySales([35, 36], {'A': [20.0, 20.0], 'N': [10.0, 10.0], 'F': [10.0, 10.0], 'B': [5.0, 5.0]})
    m = {(s.scope, s.basis): s for s in summarize(validate([ok, nograde, nofcst, b], actual, range(35, 37)))}
    assert {m[('G', x)].n_demand for x in ('가중평균', '실적기준', 'FCST기준')} == {m[('공통', '가중평균')].n_demand}
    assert {m[('G', x)].n_lt for x in ('기존수기', '가중평균', '실적기준', 'FCST기준')} == {m[('공통', '가중평균')].n_lt}
    assert m[('전체', '실적기준')].n_demand == 4  # 전체는 기준별로 채점 가능한 SKU를 모두 쓴다


def test_actual_only_sku_appended_with_memo():
    res = compute_sku(SkuInputs('X', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 1.0, 50, [10.0], [10.0] * 13))
    rows = validate([res], WeeklySales([35], {'X': [7.0], 'ONLY': [3.0]}), range(35, 36))
    assert [r.sku for r in rows] == ['X', 'ONLY']
    only = rows[-1]
    assert only.memo == '계산대상없음'
    assert all(c.error_week is None and c.covered is None for c in only.checks.values())
    m = {(s.scope, s.basis): s for s in summarize(rows)}
    assert m[('전체', '가중평균')].n_demand == 1  # 지표에서 자동으로 빠진다
    assert '미분류' not in {s.scope for s in summarize(rows)}  # 계산대상없음 행만으로 그룹 범위를 만들지 않는다


def test_error_share_sums_to_one_per_basis():
    ok, nograde, nofcst = _three_skus()
    b = compute_sku(SkuInputs('B', 'H', 'n', 'B', 0.7, 0.3, 14, 30, 1.0, 50, [10.0, 10.0], [10.0] * 13))
    actual = WeeklySales([35, 36], {'A': [20.0, 20.0], 'N': [10.0, 10.0], 'F': [10.0, 10.0], 'B': [5.0, 5.0]})
    rows = validate([ok, nograde, nofcst, b], actual, range(35, 37))
    for basis in ('가중평균', '실적기준', 'FCST기준'):
        shares = [r.checks[basis].error_share for r in rows if r.checks[basis].error_share is not None]
        assert sum(shares) == approx(1.0), basis
        total = sum(abs(r.checks[basis].error_week) for r in rows if r.checks[basis].error_week is not None)
        for r in rows:
            c = r.checks[basis]
            assert c.error_share == (None if c.error_week is None else approx(abs(c.error_week) / total))
    assert all(r.checks['기존수기'].error_share is None for r in rows)


def test_abnormal_correction_excluded_from_actual_fcst_only():
    # k=29.82(3 초과): 실적기준·FCST기준은 채점 제외, 가중평균·기존수기는 그대로 채점
    bad = compute_sku(SkuInputs('K', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 29.82, 50, [10.0] * 4, [12.0] * 13))
    ok = compute_sku(SkuInputs('O', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 1.0, 50, [10.0] * 4, [12.0] * 13))
    actual = WeeklySales(list(range(35, 44)), {'K': [20.0] * 9, 'O': [20.0] * 9})
    rows = validate([bad, ok], actual, range(35, 44))
    k = rows[0]
    for b in ('실적기준', 'FCST기준'):
        c = k.checks[b]
        assert c.excluded == '보정계수이상'
        assert (c.target_qty, c.predicted_week, c.error_week, c.covered, c.actual_days) == (None,) * 5
    assert k.checks['가중평균'].excluded is None and k.checks['가중평균'].covered is not None
    assert k.checks['기존수기'].excluded is None and k.checks['기존수기'].target_qty is not None
    assert rows[1].checks['실적기준'].excluded is None
    m = {(s.scope, s.basis): s for s in summarize(rows)}
    assert m[('공통', '가중평균')].n_demand == 1 and m[('전체', '가중평균')].n_demand == 2
    assert m[('공통', '실적기준')].n_lt == 1
    # k=0도 이상, 다른 제외 사유가 있으면 그 사유가 우선
    zero = compute_sku(SkuInputs('Z', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 0.0, 50, [10.0] * 4, [12.0] * 13))
    noact = compute_sku(SkuInputs('N', 'G', 'n', 'B', 0.7, 0.3, 14, 30, 29.82, 50, None, [12.0] * 13))
    r = validate([zero, noact], WeeklySales(list(range(35, 44)), {'Z': [1.0] * 9, 'N': [1.0] * 9}), range(35, 44))
    assert r[0].checks['FCST기준'].excluded == '보정계수이상'
    assert r[1].checks['실적기준'].excluded == '실적없음' and r[1].checks['FCST기준'].excluded == '보정계수이상'
