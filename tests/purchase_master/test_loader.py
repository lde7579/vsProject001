import pytest

from safety_stock.common.xlsx import InputError
from safety_stock.purchase_master.loader import load_purchase_master, master_columns
from safety_stock.purchase_master.model import z_for_grade

HEAD = ['품목그룹', 'SKU', '품목명', '리드타임(일)', '안전재고 값', '재고등급', '실적 가중치', 'FCST 가중치',
        '구매단가', '기초재고', '기말재고', 'FCST', '리드타임', '품질이슈']
X = [1000, 10, 20, 5, 'B', 'C']  # HEAD 뒤쪽 등급 입력 열 값


def test_load_purchase_master_reads_rows(make_wb):
    wb = make_wb({'W30 구매마스터': [HEAD,
        ['ACL', ' ACL131T0CJDG', '공기청정기', 100, 45, 'C', 0.7, 0.3, *X],
        ['FLT', 'FLTA15C1A2WH', '필터', 30, '#N/A', '#N/A', '0.9', '0.1', *X]]})
    pm = load_purchase_master(wb)
    r = pm['ACL131T0CJDG']
    assert (r.grade, r.lead_time_days, r.current_ss_days, r.weight_actual, r.weight_fcst) == ('C', 100.0, 45.0, 0.7, 0.3)
    f = pm['FLTA15C1A2WH']
    assert f.grade is None and f.current_ss_days is None and f.weight_actual == 0.9


def test_header_row_not_first(make_wb):
    wb = make_wb({'W30 구매마스터': [['재고 등급 : IF(...)'], [], [], [], [], HEAD,
                                    ['ACL', 'ACL131T0CJDG', 'x', 100, 45, 'c', 0.7, 0.3, *X]]})
    assert load_purchase_master(wb)['ACL131T0CJDG'].grade == 'C'


def test_missing_column_raises(make_wb):
    wb = make_wb({'W30 구매마스터': [[h for h in HEAD if h != 'FCST 가중치']]})
    with pytest.raises(InputError, match='FCST가중치|FCST 가중치'):
        load_purchase_master(wb)


def test_z_for_grade():
    assert z_for_grade('A') == 1.645 and z_for_grade(' b ') == 1.282 and z_for_grade('C') == 1.036
    assert z_for_grade(None) is None and z_for_grade('D') is None


def test_duplicate_sku_keeps_first(make_wb):
    wb = make_wb({'구매마스터': [HEAD,
        ['A', 'X1', 'first', 10, 5, 'A', 1, 0, *X],
        ['B', 'x1', 'second', 20, 6, 'B', 0, 1, *X]]})
    pm = load_purchase_master(wb)
    assert len(pm) == 1 and pm['X1'].name == 'first'


# 실제 시트 제목 순서(A~Z) 그대로
FULL_HEAD = ['품목그룹', 'SKU', '품목명', '구매단가', '기초재고', '기말재고', 'FCST', '회전율', 'MOQ', '구매단위',
             '리드타임(일)', '안전재고 방식', '안전재고 값', '재고등급', '실적 가중치', 'FCST 가중치', '공급업체명',
             '단가', '회전율', '리드타임', '품질이슈', '단가', '회전율', '리드타임', '품질이슈', '종합']
ACL = ['ACL', 'ACL131T0CJDG', '공기청정기', 93386, 722, 1636, 121, 0.1026, 1036, 1036, 100, '기본-일자', 45,
       'C', 0.7, 0.3, 'x', 'B', 'B', 'C', 'C', 3, 5, 1, 0, 9]


def test_master_columns_picks_by_occurrence(make_wb):
    wb = make_wb({'W30 구매마스터': [FULL_HEAD, ACL]})
    cols = master_columns(wb['W30 구매마스터'], 1)
    assert (cols['excel_turnover'], cols['turnover_grade'], cols['turnover_score']) == (8, 19, 23)
    assert (cols['price_grade'], cols['lt_grade'], cols['quality_grade']) == (18, 20, 21)
    assert (cols['price_score'], cols['lt_score'], cols['quality_score'], cols['total']) == (22, 24, 25, 26)


def test_reads_grade_inputs(make_wb):
    r = load_purchase_master(make_wb({'W30 구매마스터': [FULL_HEAD, ACL]}))['ACL131T0CJDG']
    assert (r.price, r.begin_stock, r.end_stock, r.fcst_qty) == (93386, 722, 1636, 121)
    assert (r.lt_grade, r.quality_grade, r.excel_price_grade, r.excel_turnover_grade) == ('C', 'C', 'B', 'B')
    assert (r.excel_turnover, r.excel_total, r.grade, r.sheet_row) == (0.1026, 9, 'C', 2)


def test_missing_grade_input_column_raises(make_wb):
    head = [h for h in FULL_HEAD if h != '품질이슈']
    with pytest.raises(InputError, match='품질이슈'):
        load_purchase_master(make_wb({'W30 구매마스터': [head]}))
