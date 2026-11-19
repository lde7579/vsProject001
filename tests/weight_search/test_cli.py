import openpyxl

from safety_stock.weight_search.__main__ import main

from safety_stock.inventory_days.calc import CalcSettings
from safety_stock.weight_search.search import run_search


def result_metrics(base, actual):
    return run_search(base, actual, CalcSettings(), range(35, 39), range(39, 44), range(35, 44)).metrics


SHEETS = ['후보성능', '선택가중치', '검증비교', '적용기준']


def test_cli_writes_sheets_and_prints(tmp_path, make_books, capsys):
    base, actual = make_books(tmp_path)
    out = tmp_path / '탐색.xlsx'
    assert main([str(base), '--actual', str(actual), '-o', str(out)]) == 0
    wb = openpyxl.load_workbook(out, data_only=True)
    assert wb.sheetnames == SHEETS
    text = capsys.readouterr().out
    assert '검증 구간 가중평균 WAPE:' in text and '→' in text
    heads = {n: [c.value for c in wb[n][1]] for n in SHEETS[:3]}
    assert heads['후보성능'] == ['범위', '실적가중치', 'SKU수', 'WAPE']
    assert heads['선택가중치'][:4] == ['품목그룹', 'SKU수', '현재가중치', '선택가중치']
    assert heads['검증비교'][:4] == ['구간', '범위', '가중치', '기준']
    assert wb['후보성능'].max_row == 1 + 11 * 2  # 전체 + FLT, 후보 11개씩
    rows = list(wb['검증비교'].iter_rows(min_row=2, values_only=True))
    assert {r[3] for r in rows} == {'가중평균', '기존수기'}
    assert {r[0] for r in rows} == {'검증', '전체구간'} and {r[2] for r in rows} == {'현재', '탐색'}
    assert '공통' not in {r[1] for r in rows}
    assert heads['검증비교'] == ['구간', '범위', '가중치', '기준', '수요비교SKU수', 'WAPE', '편향',
                             'LT검증SKU수', '결품방어율', '실제기준재고일수중앙값']
    # 한 행이 SummaryMetric과 같은 값인지
    m = next(m for m in result_metrics(base, actual)[('검증', '탐색')] if m.scope == '전체' and m.basis == '가중평균')
    row = next(r for r in rows if r[:4] == ('검증', '전체', '탐색', '가중평균'))
    assert row[4:] == (m.n_demand, m.wape, m.bias, m.n_lt, m.coverage, m.median_actual_days)


def test_cli_bad_weeks(tmp_path, make_books, capsys):
    base, actual = make_books(tmp_path)
    out = tmp_path / '탐색.xlsx'
    assert main([str(base), '--actual', str(actual), '-o', str(out), '--select-weeks', '38-35']) == 2
    assert '--select-weeks' in capsys.readouterr().out and not out.exists()


def test_cli_missing_file(tmp_path, capsys):
    missing = tmp_path / 'nope.xlsx'
    assert main([str(missing), '--actual', str(missing), '-o', str(tmp_path / 'o.xlsx')]) == 2
    assert '찾을 수 없습니다' in capsys.readouterr().out


def test_cli_select_start_not_after_actual_last_week(tmp_path, make_books, capsys):
    base, actual = make_books(tmp_path)
    out = tmp_path / '탐색.xlsx'
    assert main([str(base), '--actual', str(actual), '-o', str(out), '--select-weeks', '31-38']) == 2
    assert '실적 마지막 주차' in capsys.readouterr().out and not out.exists()
