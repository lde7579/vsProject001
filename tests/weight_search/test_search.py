import pytest

from safety_stock.inventory_days.calc import CalcSettings
from safety_stock.inventory_days.validate import BasisCheck, ValidationRow
from safety_stock.purchase_master.model import PurchaseMasterRow
from safety_stock.weight_search.search import (
    CANDIDATES, MIN_GROUP_SKUS, Choice, choose, current_weights, group_wape, run_search, with_weights,
)


def _pm(sku, group, w):
    return PurchaseMasterRow(sku=sku, group=group, name=None, grade='A', weight_actual=w,
                             weight_fcst=round(1 - w, 1), lead_time_days=14, current_ss_days=30)


PM = {r.sku: r for r in [_pm('g1', 'G', 0.9), _pm('g2', 'G', 0.9), _pm('g3', 'G', 0.7),
                         _pm('h1', 'H', 0.3), _pm('h2', 'H', 0.7), _pm('n1', None, 0.5)]}


def test_candidates():
    assert CANDIDATES == (0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0)
    assert MIN_GROUP_SKUS == 5


def test_current_weights_mode_and_tie():
    assert current_weights(PM) == {'G': 0.9, 'H': 0.3}


def test_with_weights_only_listed_groups():
    out = with_weights(PM, {'G': 0.4})
    assert (out['g1'].weight_actual, out['g1'].weight_fcst) == (0.4, 0.6)
    assert out['h1'] == PM['h1'] and out['n1'] == PM['n1']
    assert PM['g1'].weight_actual == 0.9  # 원본은 그대로


def _vrow(sku, group, error, actual):
    check = BasisCheck(None, None, error, None, None)
    return ValidationRow(sku, group, 14, actual, None, {'가중평균': check}, None)


def test_group_wape():
    rows = [_vrow('a', 'G', 10, 100), _vrow('b', 'G', -30, 50),
            _vrow('c', 'Z', 5, 0),            # 실제합 0 -> None
            _vrow('d', None, 2, 10),          # 그룹 없음 -> 미분류
            _vrow('e', 'G', None, 40),        # 오차 없음 -> 제외
            _vrow('f', 'G', 3, None)]         # 실제 없음 -> 제외
    out = group_wape(rows)
    assert out['G'] == (2, pytest.approx(40 / 150))
    assert out['Z'] == (1, None)
    assert out['미분류'] == (1, pytest.approx(0.2))
    assert out['전체'][0] == 4 and out['전체'][1] == pytest.approx(47 / 160)


def _table(spec):
    """spec: 그룹 -> (SKU 수, {w: wape}). 지정 안 한 w는 None."""
    return {w: {g: (n, waps.get(w)) for g, (n, waps) in spec.items()} for w in CANDIDATES}


FLAT = {w: 0.5 for w in CANDIDATES}


def test_choose_group_small_and_ties():
    spec = {
        '전체': (14, {**FLAT, 0.6: 0.1}),
        'G': (6, {**FLAT, 0.3: 0.1}),                    # w=0.3 최소 -> 그룹최적
        'S': (2, {**FLAT, 0.0: 0.01}),                   # 작은 그룹 -> 전체 최소 0.6
        'T': (6, {**FLAT, 0.2: 0.1, 0.8: 0.1 + 1e-13}),  # 동점, 현재 0.7 -> 0.8
        'V': (6, {**FLAT, 0.2: 0.1, 0.4: 0.1}),          # 동점, 거리 같음 -> 작은 w
        'U': (6, {}),                                    # 모두 None -> 현재유지
    }
    cur = {'G': 0.9, 'S': 0.1, 'T': 0.7, 'U': 0.5, 'V': 0.3}
    got = {c.group: c for c in choose(_table(spec), cur)}
    assert list(got) == list(cur)
    assert got['G'] == Choice('G', 6, 0.9, 0.3, '그룹최적')
    assert got['S'] == Choice('S', 2, 0.1, 0.6, '공통최적')
    assert got['T'].chosen == 0.8 and got['T'].method == '그룹최적'
    assert got['V'].chosen == 0.2
    assert got['U'] == Choice('U', 6, 0.5, 0.5, '현재유지')


def test_choose_pooled_all_none_keeps_current():
    spec = {'전체': (3, {}), 'S': (2, {0.3: 0.1})}
    assert choose(_table(spec), {'S': 0.7}) == [Choice('S', 2, 0.7, 0.7, '현재유지')]


def test_choose_boundary_min_group_skus():
    spec = {'전체': (9, {**FLAT, 0.6: 0.1}), 'A': (5, {**FLAT, 0.3: 0.1}), 'B': (4, {**FLAT, 0.3: 0.1})}
    got = {c.group: c for c in choose(_table(spec), {'A': 0.9, 'B': 0.9})}
    assert (got['A'].chosen, got['A'].method) == (0.3, '그룹최적')
    assert (got['B'].chosen, got['B'].method) == (0.6, '공통최적')


def test_selection_ignores_holdout(tmp_path, make_books):
    b1, a1 = make_books(tmp_path, holdout_qty=20, tag='1')
    b2, a2 = make_books(tmp_path, holdout_qty=500, tag='2')
    args = (CalcSettings(), range(35, 39), range(39, 44), range(35, 44))
    r1, r2 = run_search(b1, a1, *args), run_search(b2, a2, *args)
    assert r1.choices == r2.choices and r1.table == r2.table
    assert set(r1.metrics) == {('검증', '현재'), ('검증', '탐색'), ('전체구간', '현재'), ('전체구간', '탐색')}
    # 검증 구간 수치에는 검증 구간 데이터가 쓰인다
    assert r1.metrics[('검증', '현재')] != r2.metrics[('검증', '현재')]


def test_choose_tie_distance_float_safe():
    # 현재 0.2에서 0.1과 0.3은 거리가 같다 -> 작은 w(0.1)
    spec = {'전체': (9, FLAT), 'X': (6, {**FLAT, 0.1: 0.1, 0.3: 0.1})}
    assert choose(_table(spec), {'X': 0.2})[0].chosen == 0.1


def test_choose_small_group_own_wape_all_none_keeps_current():
    # 작은 그룹(n<5)이고 자기 WAPE가 모두 None이면 전체 값이 있어도 현재유지
    spec = {'전체': (9, {**FLAT, 0.6: 0.1}), 'S': (2, {})}
    assert choose(_table(spec), {'S': 0.7}) == [Choice('S', 2, 0.7, 0.7, '현재유지')]


def test_with_weights_skips_rows_without_weights():
    pm = {'x': _pm('x', 'G', 0.7), 'y': PurchaseMasterRow(sku='y', group='G', name=None, grade='A',
                                                           weight_actual=None, weight_fcst=None,
                                                           lead_time_days=14, current_ss_days=30)}
    out = with_weights(pm, {'G': 0.2})
    assert out['x'].weight_actual == 0.2 and out['y'] == pm['y']


def test_current_and_searched_use_same_sku_set(tmp_path, make_books):
    import openpyxl
    base, actual = make_books(tmp_path)
    wb = openpyxl.load_workbook(base)
    ws = wb['W30 구매마스터']
    ws['E3'], ws['F3'] = None, None  # BBB 가중치 없음
    wb.save(base)
    res = run_search(base, actual, CalcSettings(), range(35, 39), range(39, 44), range(35, 44))

    def n(key):
        return next(m.n_demand for m in res.metrics[key] if m.scope == '전체' and m.basis == '가중평균')
    assert n(('검증', '현재')) == n(('검증', '탐색')) == 1


def test_current_wape_filled_when_current_not_a_candidate(tmp_path, make_books):
    import openpyxl
    from safety_stock.weight_search.report import write_search_report
    base, actual = make_books(tmp_path)
    wb = openpyxl.load_workbook(base)
    ws = wb['W30 구매마스터']
    for row in (2, 3):
        ws[f'E{row}'], ws[f'F{row}'] = 0.75, 0.25  # 후보(0.1 간격)에 없는 현재 가중치
    wb.save(base)
    res = run_search(base, actual, CalcSettings(), range(35, 39), range(39, 44), range(35, 44))
    assert res.choices[0].current == 0.75
    n, wape = res.current_wape['FLT']
    assert n == 2 and wape is not None
    out = tmp_path / '탐색.xlsx'
    write_search_report(out, res, {})
    sheet = openpyxl.load_workbook(out, data_only=True)['선택가중치']
    row = next(sheet.iter_rows(min_row=2, values_only=True))
    assert row[5] == pytest.approx(wape)
