import openpyxl
import pytest

from safety_stock.__main__ import main
from safety_stock.inventory_days.report import REPORT_SHEETS


@pytest.fixture
def minimal_books(make_wb):
    """기준 엑셀과 실제 출고 엑셀을 임시 폴더에 저장하는 함수."""
    def build(tmp_path):
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
                                   ['AAA', *([20] * 9)], ['BBB', *([5] * 9)]]})
        base_path, actual_path = tmp_path / 'base.xlsx', tmp_path / 'actual.xlsx'
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
