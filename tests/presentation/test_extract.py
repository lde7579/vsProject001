"""발표 추출(extract.py)의 세 기준 분포와 사례 값 검사."""
import statistics
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / 'presentation'))

from extract import basis_spread, case_dict, safety_distribution  # noqa: E402


def test_basis_spread_values():
    s = basis_spread([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11], old=[20] * 11)
    assert s['median'] == 6
    assert s['min'] == 1
    assert s['max'] == 11
    assert s['p10'] == 2.0
    assert s['p90'] == 10.0
    assert s['distinct'] == 11
    assert s['shorter'] == 11


def _row(old, w, a, f):
    return {'기존안전재고일수': old, '가중평균_안전재고일수': w,
            '실적기준_안전재고일수': a, 'FCST기준_안전재고일수': f}


def test_safety_bases_skip_rows_with_none():
    calc = [_row(10, 8.0, 6.0, 3.0), _row(20, 12.0, 9.0, 5.0), _row(15, 9.0, 7.0, None)]
    out = safety_distribution(calc)
    b = out['bases']
    assert b['n'] == 2
    assert b['기존수기']['median'] == statistics.median([10, 20])
    assert b['가중평균']['median'] == statistics.median([8.0, 12.0])
    assert b['실적기준']['median'] == statistics.median([6.0, 9.0])
    assert b['FCST기준']['median'] == statistics.median([3.0, 5.0])
    assert b['기존수기']['shorter'] == 0
    assert b['FCST기준']['shorter'] == 2
    # 기존 키는 그대로 있어야 한다
    for k in ('n', 'old', 'new', 'old_median', 'new_median', 'labels'):
        assert k in out


def test_case_dict_has_actual_and_fcst_basis():
    r = {k: 1 for k in ['SKU', '품목그룹', '품목명', '재고등급', '실적가중치', 'FCST가중치', '리드타임(일)',
                        '현재고', '실적평균', 'FCST평균', '가중평균_수요', '가중평균_σ', '가중평균_LT평균수요',
                        '가중평균_LT수요편차', '가중평균_Z', '가중평균_안전재고', '가중평균_ROP',
                        '가중평균_적정재고일수', '가중평균_안전재고일수', '가중평균_현재고일수',
                        '가중평균_발주', '기존안전재고일수', '기존수기_적정재고일수']}
    r.update({'실적기준_안전재고일수': 5.45, '실적기준_적정재고일수': 9.9,
              'FCST기준_안전재고일수': 4.01, 'FCST기준_적정재고일수': 7.7})
    v = {k: 1 for k in ['실제주평균', '실제LT수요', '기존수기_목표재고량', '가중평균_목표재고량',
                        '기존수기_LT수요방어', '가중평균_LT수요방어', '기존수기_실제기준재고일수',
                        '가중평균_실제기준재고일수', '가중평균_수요오차']}
    v.update({'실적기준_LT수요방어': '방어', '실적기준_실제기준재고일수': 12.5,
              'FCST기준_LT수요방어': '결품', 'FCST기준_실제기준재고일수': 3.5})
    c = case_dict(r, v, 'rule')
    assert (c['act_s'], c['fc_s']) == (5.45, 4.01)
    assert (c['act_u'], c['fc_u']) == (9.9, 7.7)
    assert (c['act_cover'], c['fc_cover']) == ('방어', '결품')
    assert (c['act_actual_days'], c['fc_actual_days']) == (12.5, 3.5)
