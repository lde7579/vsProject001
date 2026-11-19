from safety_stock.virtual_data.__main__ import main


def test_cli_rejects_non_xlsx_file(tmp_path, capsys):
    bad = tmp_path / 'bad.xlsx'
    bad.write_text('엑셀이 아닌 파일', encoding='utf-8')
    out = tmp_path / 'o.xlsx'
    assert main([str(bad), '-o', str(out)]) == 2
    text = capsys.readouterr().out
    assert '입력 오류' in text and '엑셀 파일을 열 수 없습니다' in text and not out.exists()
