import pytest

from safety_stock.common.xlsx import InputError
from safety_stock.inventory_days.loader import load_onhand, load_sales, load_summary


def test_load_sales_keeps_blank_as_none(make_wb):
    wb = make_wb({'Sales_WMS 출고': [['SKU', 'SKU명', 'WK01', 'WK02', None, '6주 평균'],
                                     ['HWCM1401K0WH', '카트리지', None, 131, None, 50]]})
    s = load_sales(wb)
    assert s.week_numbers == [1, 2] and s.by_sku['HWCM1401K0WH'] == [None, 131.0]


def test_load_sales_falls_back_to_upload_sheet(make_wb):
    wb = make_wb({'출고실적': [['SKU', 'SKU명', 'WK35'], ['ACL131T0CJDG', 'x', 10]]})
    assert load_sales(wb).by_sku == {'ACL131T0CJDG': [10.0]}


def test_load_sales_first_row_wins_and_normalizes_sku(make_wb):
    wb = make_wb({'출고실적': [['SKU', 'WK01'], [' acl1 ', 1], ['ACL1', 2]]})
    assert load_sales(wb).by_sku == {'ACL1': [1.0]}


def test_load_onhand_sums_warehouses(make_wb):
    wb = make_wb({'Summary_가중평균수요기준': [['SKU', '현재고']],
                  '0824_01 현재고': [['창고', '품목 ID', '현 재고'], ['P100', 'ACL131T0CJDG', 635], ['S940', 'ACL131T0CJDG', '87']]})
    assert load_onhand(wb) == {'ACL131T0CJDG': 722.0}


def test_load_onhand_missing_sheet_raises(make_wb):
    wb = make_wb({'Summary_현재고': [['SKU', '현재고']]})
    with pytest.raises(InputError):
        load_onhand(wb)


def test_load_summary_first_duplicate_header_and_demand_column(make_wb):
    wb = make_wb({'Summary_가중평균수요기준': [
        [None] * 8 + ['가중평균', '실적기준'],
        ['SKU그룹', 'SKU', '현재고', '가중평균 수요', 'LT 평균수요', 'ROP (재주문점)', '과소예측\n보정계수', '재고등급', '적정\n재고일수', '적정\n재고일수'],
        ['ACL', 'ACL131T0CJDG', 722, 134, 1914, '2,109', 1.33, 'C', 110, 108]]})
    row = load_summary(wb, 'Summary_가중평균')['ACL131T0CJDG']
    assert row['demand'] == 134 and row['lt_demand'] == 1914 and row['rop'] == 2109
    assert row['target_days'] == 110 and row['correction'] == 1.33 and row['grade'] == 'C' and row['group'] == 'ACL'


def test_load_summary_missing_sheet_raises(make_wb):
    wb = make_wb({'x': [['SKU']]})
    with pytest.raises(InputError):
        load_summary(wb, 'Summary_FCST')


def test_load_summary_text_values_and_missing_columns(make_wb):
    wb = make_wb({'Summary_FCST기준': [
        ['SKU', '향후 13주 평균수요', '적정재고일수'],
        ['A1', '#DIV/0!', '(13)']]})
    row = load_summary(wb, 'Summary_FCST')['A1']
    assert row['demand'] is None and row['target_days'] == -13 and row['grade'] is None


def test_load_summary_required_column_missing_raises(make_wb):
    wb = make_wb({'Summary_FCST기준': [['SKU', '현재고'], ['A1', 5]]})
    with pytest.raises(InputError, match='적정재고일수'):
        load_summary(wb, 'Summary_FCST', required=('target_days',))
    assert load_summary(wb, 'Summary_FCST')['A1']['onhand'] == 5
