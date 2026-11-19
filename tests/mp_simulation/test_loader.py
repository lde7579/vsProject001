from safety_stock.mp_simulation.loader import first_weeks, load_weekly_plan

ROWS = [
    ['평가클래스', '품목그룹', '품목 ID', '품목명', '구분', '2026', None, None],
    [None, None, None, None, None, 'W35(08/24)', 'W36(08/31)', 'W37(09/07)'],
    [None, 'TOTAL', None, None, 'SCM FCST', 300, 300, 300],
    ['7940', 'ACL', 'ACL131T0CJDG', '공기청정기', 'SCM FCST', 121, None, '105'],
    ['7940', 'ACL', 'ACL131T0CJDG', '중복행', 'SCM FCST', 999, 999, 999],
]


def test_load_weekly_plan_two_row_header(make_wb):
    plan = load_weekly_plan(make_wb({'0824_02 SCM FCST': ROWS}), 'SCM FCST')
    assert plan.week_labels == ['W35(08/24)', 'W36(08/31)', 'W37(09/07)']
    assert plan.by_sku == {'ACL131T0CJDG': [121.0, None, 105.0]}


def test_first_weeks_fills_blank_with_zero(make_wb):
    plan = load_weekly_plan(make_wb({'0824_02 SCM FCST': ROWS}), 'SCM FCST')
    assert first_weeks(plan, 2) == {'ACL131T0CJDG': [121.0, 0.0]}
    assert first_weeks(plan, 13) == {'ACL131T0CJDG': [121.0, 0.0, 105.0]}


def test_week_order_is_column_order_across_year_boundary(make_wb):
    rows = [
        ['품목 ID', '2026', None, '2027'],
        [None, 'W52(12/21)', 'W53(12/28)', 'W01(01/04)'],
        ['sku1 ', 1, 2, 3],
    ]
    plan = load_weekly_plan(make_wb({'기말재고': rows}), '기말재고')
    assert plan.week_labels == ['W52(12/21)', 'W53(12/28)', 'W01(01/04)']
    assert first_weeks(plan, 3) == {'SKU1': [1.0, 2.0, 3.0]}
