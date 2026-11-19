import pytest

from safety_stock.common.xlsx import (
    InputError,
    find_header_row,
    find_sheet,
    header_cols,
    header_map,
    norm_header,
    norm_sku,
    require_cols,
    to_number,
)


def test_to_number_converts_display_strings():
    assert to_number(3) == 3.0
    assert to_number('1,036') == 1036.0
    assert to_number('10.26%') == pytest.approx(0.1026)
    assert to_number('(13)') == -13.0
    for v in ['', '  ', '#N/A', '#DIV/0!', 'ORDER', None]:
        assert to_number(v) is None


def test_to_number_rejects_non_finite():
    for v in ['nan', 'NaN', 'inf', '-inf', 'Infinity', float('nan'), float('inf'), float('-inf')]:
        assert to_number(v) is None, v


def test_norm_sku_and_header():
    assert norm_sku(' acl131t0cjdg ') == 'ACL131T0CJDG'
    assert norm_sku('') is None and norm_sku(None) is None
    assert norm_header('기존\n안전재고일수') == '기존안전재고일수'
    assert norm_header('ROP (재주문점)') == 'ROP(재주문점)'


def test_find_sheet_by_keyword(make_wb):
    wb = make_wb({'0824_02 SCM FCST': [['x']], 'W30 구매마스터': [['y']]})
    assert find_sheet(wb, 'SCM FCST').title == '0824_02 SCM FCST'
    with pytest.raises(InputError, match='구매마스터|없'):
        find_sheet(wb, '현재고')


def test_find_header_row_skips_description_rows(make_wb):
    wb = make_wb({'s': [['재고 등급 : IF(...)'], [], [], [], [None, 'MP결과'], ['품목그룹', 'SKU', '품목명']]})
    ws = wb['s']
    assert find_header_row(ws, 'SKU') == 6
    assert header_map(ws, 6)['SKU'] == 2
    with pytest.raises(InputError, match='재고등급'):
        require_cols(header_map(ws, 6), ['SKU', '재고등급'], 's')


def test_header_cols_lists_duplicates(make_wb):
    wb = make_wb({'s': [['단가', '회전율', ' 단가', '회전율', '회전 율']]})
    assert header_cols(wb['s'], 1) == {'단가': [1, 3], '회전율': [2, 4, 5]}
