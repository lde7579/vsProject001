import pytest

from safety_stock.purchase_master.grade import (
    compute_grades, grade_from_rank, grade_verdict, percentrank_inc, turnover)
from safety_stock.purchase_master.model import PurchaseMasterRow


def _row(sku, price=None, b=None, e=None, f=None, lt=None, q=None, **excel):
    """등급 산출 테스트용 PurchaseMasterRow 생성 도우미."""
    return PurchaseMasterRow(sku, None, None, excel.pop('grade', None), None, None, None, None,
                             price=price, begin_stock=b, end_stock=e, fcst_qty=f,
                             lt_grade=lt, quality_grade=q, **excel)


def test_percentrank_ties():
    assert percentrank_inc([1, 2, 2, 3], 2) == 0.333
    assert percentrank_inc([1, 2, 2, 3], 3) == 1.0 and percentrank_inc([5], 5) == 1.0


def test_percentrank_truncates():
    vals = list(range(1, 102))            # n=101, x=27 → 26/100 = 0.26
    assert percentrank_inc(vals, 27) == 0.26 and grade_from_rank(0.26) == 'B'
    assert percentrank_inc(list(range(1, 8)), 3) == 0.333   # 2/6=0.3333… 내림
    assert grade_from_rank(0.259) == 'C' and grade_from_rank(0.75) == 'A'


def test_turnover():
    assert turnover(121, 722, 1636) == pytest.approx(121 / 1179)
    assert turnover(10, 0, 0) is None and turnover(None, 5, 5) == 0.0


def test_compute_grades_scores_and_zero_turnover():
    rows = [_row('A', price=100, b=10, e=10, f=50, lt='A', q='A'),
            _row('B', price=50, b=10, e=10, f=5, lt='B', q='C'),
            _row('C', price=10, b=0, e=0, f=5, lt='C', q='B')]
    g = compute_grades(rows)
    assert list(g) == ['A', 'B', 'C']
    assert (g['C'].turnover, g['C'].turnover_grade) == (None, 'C')
    assert g['A'].scores == (5, 10, 5, 5) and g['A'].total == 25 and g['A'].grade == 'A'
    assert g['B'].scores == (3, 1, 3, 1) and g['B'].total == 8   # 품절이력 C = 1점 · 회전율 배열은 A,B 두 개 → B 0.0 → C
    assert g['C'].grade == 'C'


def test_ungradable_row_has_memo_and_is_excluded():
    rows = [_row('A', price=None, lt='A', q='X'), _row('B', price=1, lt='A', q='A'), _row('C', price=2, lt='B', q='B')]
    g = compute_grades(rows)
    assert g['A'].grade is None and g['A'].memo == '등급산출불가(구매단가없음, 품절이력등급없음)'
    assert g['A'].scores is None and g['A'].total is None and g['A'].price_rank is None
    assert g['B'].price_rank == 0.0 and g['C'].price_rank == 1.0    # A는 단가 배열에서 빠짐


def test_grade_verdict():
    rows = [_row('A', price=100, b=10, e=10, f=50, lt='A', q='A'),
            _row('B', price=50, b=10, e=10, f=5, lt='B', q='C')]
    g = compute_grades(rows)
    ra = _row('A', price=100, b=10, e=10, f=50, lt='A', q='A', excel_price_grade='A',
              excel_turnover_grade='A', excel_total=25.0, grade='A')
    assert grade_verdict(ra, g['A']) == '일치'
    rb = _row('A', price=100, b=10, e=10, f=50, lt='A', q='A', excel_price_grade='A',
              excel_turnover_grade='A', excel_total=24.0, grade='A')
    assert grade_verdict(rb, g['A']) == '불일치'
    assert grade_verdict(rows[1], g['B']) == '비교불가'


def test_grade_verdict_ungradable_with_excel_grade_is_mismatch():
    r = _row('A', price=None, lt='A', q='A', grade='A')
    g = compute_grades([r, _row('B', price=1, lt='A', q='A')])['A']
    assert g.grade is None and grade_verdict(r, g) == '불일치'
