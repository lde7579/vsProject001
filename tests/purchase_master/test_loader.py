import pytest

from safety_stock.common.xlsx import InputError
from safety_stock.purchase_master.loader import load_purchase_master
from safety_stock.purchase_master.model import z_for_grade

HEAD = ['품목그룹', 'SKU', '품목명', '리드타임(일)', '안전재고 값', '재고등급', '실적 가중치', 'FCST 가중치']


def test_load_purchase_master_reads_rows(make_wb):
    wb = make_wb({'W30 구매마스터': [HEAD,
        ['ACL', ' ACL131T0CJDG', '공기청정기', 100, 45, 'C', 0.7, 0.3],
        ['FLT', 'FLTA15C1A2WH', '필터', 30, '#N/A', '#N/A', '0.9', '0.1']]})
    pm = load_purchase_master(wb)
    r = pm['ACL131T0CJDG']
    assert (r.grade, r.lead_time_days, r.current_ss_days, r.weight_actual, r.weight_fcst) == ('C', 100.0, 45.0, 0.7, 0.3)
    f = pm['FLTA15C1A2WH']
    assert f.grade is None and f.current_ss_days is None and f.weight_actual == 0.9


def test_header_row_not_first(make_wb):
    wb = make_wb({'W30 구매마스터': [['재고 등급 : IF(...)'], [], [], [], [], HEAD,
                                    ['ACL', 'ACL131T0CJDG', 'x', 100, 45, 'c', 0.7, 0.3]]})
    assert load_purchase_master(wb)['ACL131T0CJDG'].grade == 'C'


def test_missing_column_raises(make_wb):
    wb = make_wb({'W30 구매마스터': [HEAD[:-1]]})
    with pytest.raises(InputError, match='FCST가중치|FCST 가중치'):
        load_purchase_master(wb)


def test_z_for_grade():
    assert z_for_grade('A') == 1.645 and z_for_grade(' b ') == 1.282 and z_for_grade('C') == 1.036
    assert z_for_grade(None) is None and z_for_grade('D') is None


def test_duplicate_sku_keeps_first(make_wb):
    wb = make_wb({'구매마스터': [HEAD,
        ['A', 'X1', 'first', 10, 5, 'A', 1, 0],
        ['B', 'x1', 'second', 20, 6, 'B', 0, 1]]})
    pm = load_purchase_master(wb)
    assert len(pm) == 1 and pm['X1'].name == 'first'
