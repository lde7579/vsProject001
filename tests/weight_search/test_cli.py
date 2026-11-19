import openpyxl

from safety_stock.weight_search.__main__ import main

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


def test_cli_bad_weeks(tmp_path, make_books, capsys):
    base, actual = make_books(tmp_path)
    out = tmp_path / '탐색.xlsx'
    assert main([str(base), '--actual', str(actual), '-o', str(out), '--select-weeks', '38-35']) == 2
    assert '--select-weeks' in capsys.readouterr().out and not out.exists()


def test_cli_missing_file(tmp_path, capsys):
    missing = tmp_path / 'nope.xlsx'
    assert main([str(missing), '--actual', str(missing), '-o', str(tmp_path / 'o.xlsx')]) == 2
    assert '찾을 수 없습니다' in capsys.readouterr().out
