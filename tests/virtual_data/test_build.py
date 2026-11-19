import openpyxl
import pytest

from safety_stock.common.xlsx import InputError
from safety_stock.inventory_days.calc import CalcSettings, build_inputs, compute_sku
from safety_stock.inventory_days.compare import compare_with_excel
from safety_stock.inventory_days.loader import load_onhand, load_sales, load_summary
from safety_stock.mp_simulation.loader import load_weekly_plan
from safety_stock.purchase_master.grade import compute_grades
from safety_stock.purchase_master.loader import load_purchase_master
from safety_stock.purchase_master.model import PurchaseMasterRow
from safety_stock.virtual_data.build import build_virtual, fill_inventory_gaps


def _row(sku, begin, end, fcst):
    return PurchaseMasterRow(sku=sku, group=None, name=None, grade=None, weight_actual=None,
                             weight_fcst=None, lead_time_days=None, current_ss_days=None,
                             begin_stock=begin, end_stock=end, fcst_qty=fcst)


def test_fill_inventory_gaps():
    rows = [_row('R1', 100, 100, 10), _row('R2', 100, 100, 30), _row('Z1', 0, 0, 50), _row('Z0', None, 0, 0)]
    assert fill_inventory_gaps(rows) == {'Z1': 250, 'Z0': 1}


@pytest.fixture
def small_book(make_wb):
    def build(tmp_path):
        weeks = [f'WK{n:02d}' for n in range(1, 32)]
        fcst_labels = [f'W{n}(08/24)' for n in range(35, 48)]
        head = ['SKU', '평균수요', '과소예측보정계수', '재고등급', 'Z값계수', '현재고', '수요표준편차', 'LT평균수요',
                'LT수요편차', '적정안전재고', 'ROP(재주문점)', '적정재고일수', '적정안전재고일수']
        master_head = ['SKU', '품목그룹', '품목명', '재고등급', '실적 가중치', 'FCST 가중치', '리드타임(일)',
                       '안전재고 값', '구매단가', '기초재고', '기말재고', 'FCST', '리드타임', '품질이슈']
        wb = make_wb({
            'W30 구매마스터': [
                master_head,
                ['AAA', 'FLT', 'a', 'C', 0.7, 0.3, 14, 30, 2000, 100, 100, 100, 'A', 'A'],
                ['BBB', 'FLT', 'b', 'C', 0.7, 0.3, 14, 30, 1000, 0, 0, 50, 'C', 'C'],
                ['CCC', 'FLT', 'c', 'C', 0.7, 0.3, 14, 30, 500, 100, 100, 20, 'B', 'B'],
            ],
            'Sales_WMS 출고': [['SKU', *weeks], ['AAA', *range(10, 41)], ['BBB', *([5] * 31)], ['CCC', *([3] * 31)]],
            '0824_02 SCM FCST': [['품목 ID', 'x', *fcst_labels], ['AAA', None, *([20] * 13)],
                                 ['BBB', None, *([4] * 13)], ['CCC', None, *([2] * 13)]],
            '0824_01 현재고': [['품목 ID', '현재고'], ['AAA', 100], ['BBB', 50], ['CCC', 10]],
            'Summary_가중평균수요기준': [head, ['AAA', 999, 1.2, 'C', 1.0], ['BBB', None, 1.2], ['CCC', 1, 1.2]],
            'Summary_실적기준': [head, ['AAA', 1, 1.2]],
            'Summary_FCST기준': [head, ['AAA', 1, 1.2]],
        })
        src = tmp_path / 'src.xlsx'
        wb.save(src)
        return src, tmp_path / 'out.xlsx'
    return build


def _compare(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    pm = load_purchase_master(wb)
    grades = {s: g.grade for s, g in compute_grades(list(pm.values())).items()}
    weighted = load_summary(wb, 'Summary_가중평균', required=('correction',))
    summaries = {'가중평균': weighted, '실적기준': load_summary(wb, 'Summary_실적'),
                 'FCST기준': load_summary(wb, 'Summary_FCST')}
    corr = {s: r.get('correction') for s, r in weighted.items()}
    inputs = build_inputs(pm, load_sales(wb), load_weekly_plan(wb, 'SCM FCST'), load_onhand(wb),
                          corr, CalcSettings(), grades=grades)
    return compare_with_excel([compute_sku(i) for i in inputs], summaries)


def test_build_virtual_makes_excel_agree(tmp_path, small_book):
    src, out = small_book(tmp_path)
    changes = build_virtual(src, out)
    wb = openpyxl.load_workbook(out, data_only=True)
    assert '가상보완_내역' in wb.sheetnames and wb['가상보완_내역'].max_row == len(changes) + 1
    assert [c.value for c in wb['가상보완_내역'][1]] == ['시트', '셀', '구분', '원래 값', '새 값', '사유']
    assert {r.status for r in _compare(out)} <= {'일치', '계산값없음'}
    reasons = {c.reason.split('(')[0] for c in changes}
    assert {'회전율 분모 0 → 가상 재고', '재고등급 산식 재계산', '산식 결과로 교체'} <= reasons
    # BBB 기초·기말재고가 채워졌다
    master = wb['W30 구매마스터']
    assert master['J3'].value == master['K3'].value and master['J3'].value > 0
    # 원본은 그대로
    orig = openpyxl.load_workbook(src)
    assert orig['Summary_가중평균수요기준']['B2'].value == 999
    assert orig['W30 구매마스터']['J3'].value == 0
    assert '가상보완_내역' not in orig.sheetnames


def test_refuses_formula_workbook(tmp_path, make_wb):
    src = tmp_path / 'f.xlsx'
    make_wb({'W30 구매마스터': [['x'], ['=1+1']]}).save(src)
    with pytest.raises(InputError, match='수식 셀'):
        build_virtual(src, tmp_path / 'o.xlsx')


def test_refuses_same_path(tmp_path, small_book):
    src, _ = small_book(tmp_path)
    with pytest.raises(InputError):
        build_virtual(src, src)


def _log(path):
    ws = openpyxl.load_workbook(path, data_only=True)['가상보완_내역']
    return [tuple(c.value for c in row) for row in ws.iter_rows(min_row=2)]


def test_every_changed_cell_is_logged(tmp_path, small_book):
    src, out = small_book(tmp_path)
    wb = openpyxl.load_workbook(src)
    wb['W30 구매마스터']['E4'], wb['W30 구매마스터']['F4'] = None, None  # CCC 가중치를 지워 계산값이 없게 한다(엑셀 값은 있음 → 기록만)
    wb.save(src)
    changes = build_virtual(src, out)
    log = _log(out)
    assert len(log) == len(changes)
    a, b = openpyxl.load_workbook(src, data_only=True), openpyxl.load_workbook(out, data_only=True)
    logged = {(sheet, cell): (kind, new) for sheet, cell, kind, _old, new, _why in log}
    diff = set()
    for ws in a.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if b[ws.title][c.coordinate].value != c.value:
                    diff.add((ws.title, c.coordinate))
    # 값이 바뀐 모든 칸은 구분=변경으로 기록되고 새 값이 결과 칸 값과 같다
    for sheet, cell in diff:
        kind, new = logged[(sheet, cell)]
        assert kind == '변경' and new == b[sheet][cell].value
    # 결과에서 안 바뀐 기록은 구분=기록만 뿐이다
    for key, (kind, new) in logged.items():
        if key not in diff:
            assert kind == '기록만' and new is None
    assert any(k == '기록만' for k, _ in logged.values())


def test_population_change_reason_for_unfilled_rows(tmp_path, make_wb):
    """가상 재고로 채운 행 때문에 다른 행의 회전율 순위가 바뀌면 사유를 따로 남긴다."""
    head = ['SKU', '품목그룹', '품목명', '재고등급', '실적 가중치', 'FCST 가중치', '리드타임(일)', '안전재고 값',
            '구매단가', '기초재고', '기말재고', 'FCST', '리드타임', '품질이슈']
    spec = [('L', 300, 100, 'B', 'C'), ('M', 100, 60, 'A', 'C'), ('Q', 100, 20, 'A', 'B'),
            ('N', 400, 20, 'B', 'C'), ('Z', 500, 1000, 'A', 'A')]  # Z만 재고 0 → 가상 재고로 채워진다
    skus = [k for k, *_ in spec]
    weeks = [f'WK{n:02d}' for n in range(1, 32)]
    sh = ['SKU', '평균수요', '과소예측보정계수', '재고등급', 'Z값계수']
    master = [head, *[[k, 'F', k, 'C', 0.7, 0.3, 14, 30, price, 0 if k == 'Z' else 10, 0 if k == 'Z' else 10,
                       fcst, lt, q] for k, price, fcst, lt, q in spec]]
    wb = make_wb({
        'Sales_WMS 출고': [['SKU', *weeks], *[[k, *([5] * 31)] for k in skus]],
        '0824_02 SCM FCST': [['품목 ID', 'x', *[f'W{n}(08/24)' for n in range(35, 48)]],
                             *[[k, None, *([4] * 13)] for k in skus]],
        '0824_01 현재고': [['품목 ID', '현재고'], *[[k, 10] for k in skus]],
        'Summary_가중평균수요기준': [sh, *[[k, 1, 1.2] for k in skus]],
        'Summary_실적기준': [sh, ['L', 1, 1.2]],
        'Summary_FCST기준': [sh, ['L', 1, 1.2]],
        'W30 구매마스터': master,
    })
    ws = wb['W30 구매마스터']
    before = compute_grades(list(load_purchase_master(wb).values()))
    for k, row in zip(skus, range(2, 7)):
        ws.cell(row=row, column=4, value=before[k].grade or 'C')  # 채우기 전 산식과 맞는 재고등급
    src = tmp_path / 's.xlsx'
    wb.save(src)
    changes = build_virtual(src, tmp_path / 'o.xlsx')
    moved = [c for c in changes if c.sheet == 'W30 구매마스터' and c.reason == '가상 재고로 회전율 순위 모집단 변경']
    assert moved and all(c.cell[1:] != '6' for c in moved)  # 채운 행(6행)은 이 사유가 아니다
    assert all(c.reason != '가상 재고로 회전율 순위 모집단 변경' for c in changes if c.cell == 'J6')
