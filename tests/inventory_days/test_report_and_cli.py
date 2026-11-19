import openpyxl
import pytest

from safety_stock.__main__ import main
from safety_stock.inventory_days.report import REPORT_SHEETS


@pytest.fixture
def minimal_books(make_wb):
    """기준 엑셀과 실제 출고 엑셀을 임시 폴더에 저장하는 함수."""
    def build(tmp_path, actual_qty=20, tag=''):
        weeks = [f'WK{n:02d}' for n in range(1, 32)]
        fcst_labels = [f'W{n}(08/24)' for n in range(35, 48)]
        summary_head = ['SKU', '평균수요', '과소예측보정계수']
        base = make_wb({
            'W30 구매마스터': [
                ['SKU', '품목그룹', '품목명', '재고등급', '실적 가중치', 'FCST 가중치', '리드타임(일)', '안전재고 값'],
                ['AAA', 'FLT', '필터', 'A', 0.7, 0.3, 14, 30],
                ['BBB', 'FLT', '필터2', 'B', 0.7, 0.3, 14, 30],
            ],
            'Sales_WMS 출고': [['SKU', *weeks], ['AAA', *range(10, 41)], ['BBB', *([5] * 31)]],
            '0824_02 SCM FCST': [['품목 ID', 'x', *fcst_labels], ['AAA', None, *([20] * 13)]],
            '0824_01 현재고': [['품목 ID', '현재고'], ['AAA', 100], ['BBB', 50]],
            'Summary_가중평균수요기준': [summary_head, ['AAA', 1, 1.2], ['BBB', 1, 1.2]],
            'Summary_실적기준': [summary_head, ['AAA', 1, 1.2]],
            'Summary_FCST기준': [summary_head, ['AAA', 1, 1.2]],
        })
        actual = make_wb({'출고실적': [['SKU', *[f'WK{n}' for n in range(35, 44)]],
                                   ['AAA', *([actual_qty] * 9)], ['BBB', *([5] * 9)]]})
        base_path, actual_path = tmp_path / f'base{tag}.xlsx', tmp_path / f'actual{tag}.xlsx'
        base.save(base_path)
        actual.save(actual_path)
        return base_path, actual_path
    return build


def test_cli_writes_all_sheets(tmp_path, minimal_books, capsys):
    base, actual = minimal_books(tmp_path)
    out = tmp_path / '결과.xlsx'
    rc = main([str(base), '--actual', str(actual), '-o', str(out)])
    assert rc == 0
    wb = openpyxl.load_workbook(out, data_only=True)
    assert wb.sheetnames == REPORT_SHEETS
    assert wb['산출결과'].max_row >= 2 and wb['검증요약']['B1'].value == '기준'
    assert wb['산출결과']['A2'].value == 'AAA'
    # 엑셀대조에는 필드별 행과 SKU 판정 열이 있다
    head = [c.value for c in wb['엑셀대조'][1]]
    assert head[-1] == 'SKU판정' and wb['엑셀대조'].max_row > 1
    # 검증_SKU에 제외 사유 열, 검증요약에 공통 행
    assert any(str(h).endswith('제외사유') for h in (c.value for c in wb['검증_SKU'][1]))
    assert '공통' in {r[0].value for r in wb['검증요약'].iter_rows(min_row=2)}
    out_text = capsys.readouterr().out
    assert '처리 SKU 수: 2' in out_text and str(out) in out_text


def test_cli_input_error_exit_code(tmp_path, capsys):
    bad = tmp_path / 'bad.xlsx'
    openpyxl.Workbook().save(bad)
    assert main([str(bad), '--actual', str(bad), '-o', str(tmp_path / 'o.xlsx')]) == 2
    assert '시트' in capsys.readouterr().out


def test_cli_missing_file_exit_code(tmp_path, capsys):
    missing = tmp_path / 'nope.xlsx'
    assert main([str(missing), '--actual', str(missing), '-o', str(tmp_path / 'o.xlsx')]) == 2
    assert '찾을 수 없습니다' in capsys.readouterr().out


def _run(tmp_path, minimal_books, **kw):
    base, actual = minimal_books(tmp_path, **kw)
    out = tmp_path / f"out{kw.get('tag', '')}.xlsx"
    assert main([str(base), '--actual', str(actual), '-o', str(out)]) == 0
    return openpyxl.load_workbook(out, data_only=True)


def _col(ws, title):
    head = [c.value for c in ws[1]]
    return head.index(title) + 1


def test_report_values_and_formats(tmp_path, minimal_books):
    wb = _run(tmp_path, minimal_books)
    ws = wb['산출결과']
    # AAA: 실적 10..40 평균 25, FCST 20 -> 25*0.7 + 20*0.3 = 23.5
    cell = ws.cell(row=2, column=_col(ws, '가중평균_수요'))
    assert cell.value == pytest.approx(23.5) and cell.number_format == '#,##0'
    assert ws.cell(row=2, column=_col(ws, '가중평균_적정재고일수')).number_format == '0.0'
    summary = wb['검증요약']
    wape_col = _col(summary, 'WAPE')
    rows = {(r[0].value, r[1].value): r for r in summary.iter_rows(min_row=2)}
    assert rows[('전체', '가중평균')][wape_col - 1].number_format == '0.0%'
    # AAA 예측 23.5, 실제 20 -> WAPE 17.5%
    assert rows[('전체', '가중평균')][wape_col - 1].value == pytest.approx(0.175)
    # 기존수기는 예측 수요가 없어 WAPE가 빈 칸
    assert rows[('전체', '기존수기')][wape_col - 1].value is None


def test_all_sheets_freeze_and_filter(tmp_path, minimal_books):
    wb = _run(tmp_path, minimal_books)
    for ws in wb:
        assert ws.freeze_panes == 'A2', ws.title
        assert ws.auto_filter.ref, ws.title


def test_actual_file_does_not_change_calculation(tmp_path, minimal_books):
    a = _run(tmp_path, minimal_books, actual_qty=20, tag='1')
    b = _run(tmp_path, minimal_books, actual_qty=999, tag='2')
    for name in ('산출결과', '엑셀대조'):
        va = [[c.value for c in r] for r in a[name].iter_rows()]
        vb = [[c.value for c in r] for r in b[name].iter_rows()]
        assert va == vb, name
    assert [[c.value for c in r] for r in a['검증요약'].iter_rows()] !=         [[c.value for c in r] for r in b['검증요약'].iter_rows()]


def test_display_labels(tmp_path, minimal_books):
    wb = _run(tmp_path, minimal_books)
    fields = {r[2].value for r in wb['엑셀대조'].iter_rows(min_row=2)}
    assert fields <= {'수요', 'σ', 'LT평균수요', 'LT수요편차', '안전재고', 'ROP', '적정재고일수', '안전재고일수', '현재고'}
    assert {'수요', '현재고'} <= fields
    ws = wb['산출결과']
    orders = {ws.cell(row=r, column=_col(ws, '가중평균_발주')).value for r in range(2, ws.max_row + 1)}
    assert orders <= {'발주필요', '정상', None} and orders & {'발주필요', '정상'}
    vs = wb['검증_SKU']
    col = _col(vs, '가중평균_LT수요방어')
    assert {vs.cell(row=r, column=col).value for r in range(2, vs.max_row + 1)} <= {'방어', '결품', None}
    assert vs.cell(row=2, column=col).value == '방어'


def test_meta_has_formulas(tmp_path, minimal_books):
    wb = _run(tmp_path, minimal_books)
    text = ' '.join(f'{r[0].value} {r[1].value}' for r in wb['적용기준'].iter_rows(min_row=2))
    for word in ('STDEV.P', '빈 칸', '발주', '주수요/7'):
        assert word in text


def test_output_same_as_input_rejected(tmp_path, minimal_books, capsys):
    base, actual = minimal_books(tmp_path)
    assert main([str(base), '--actual', str(actual), '-o', str(actual)]) == 2
    assert '같' in capsys.readouterr().out
    assert actual.exists() and openpyxl.load_workbook(actual).sheetnames == ['출고실적']


def test_save_permission_error(tmp_path, minimal_books, capsys, monkeypatch):
    base, actual = minimal_books(tmp_path)

    def boom(self, filename):
        raise PermissionError(filename)
    monkeypatch.setattr(openpyxl.Workbook, 'save', boom)
    assert main([str(base), '--actual', str(actual), '-o', str(tmp_path / 'o.xlsx')]) == 2
    assert '저장' in capsys.readouterr().out
