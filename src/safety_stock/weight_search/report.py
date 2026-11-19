"""가중치 탐색 결과 xlsx 작성."""
from pathlib import Path

import openpyxl

from safety_stock.common.xlsx import InputError
from safety_stock.inventory_days.report import PCT, SUMMARY_COLS, write_table
from safety_stock.weight_search.search import POOLED, SearchResult

SEARCH_SHEETS = ['후보성능', '선택가중치', '검증비교', '적용기준']
_SHOWN_BASES = ('가중평균', '기존수기')


def write_search_report(path, result: SearchResult, meta: dict[str, str]) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheets = {name: wb.create_sheet(name) for name in SEARCH_SHEETS}

    # 후보성능: 범위('전체'가 먼저) x 후보 w
    keys = [POOLED] + [c.group for c in result.choices]
    keys += sorted({k for t in result.table.values() for k in t} - set(keys))
    rows = [[k, w, *result.table[w].get(k, (0, None))] for k in keys for w in result.table]
    write_table(sheets['후보성능'], ['범위', '실적가중치', 'SKU수', 'WAPE'], rows, [None, '0.0', '0', PCT])

    # 선택가중치: 현재 가중치 그대로와 고른 w의 선택 구간 WAPE
    def chosen_wape(w, group):
        return result.table[w].get(group, (0, None))[1] if w in result.table else None
    rows = [[c.group, c.n, c.current, c.chosen, c.method, result.current_wape.get(c.group, (0, None))[1],
             chosen_wape(c.chosen, c.group)]
            for c in result.choices]
    write_table(sheets['선택가중치'],
                 ['품목그룹', 'SKU수', '현재가중치', '선택가중치', '선택방식', '현재WAPE(선택구간)', '선택WAPE(선택구간)'],
                 rows, [None, '0', '0.0', '0.0', None, PCT, PCT])

    # 검증비교: 기준 가중평균·기존수기, 범위 전체와 그룹('공통' 제외)
    cols = [c for c in SUMMARY_COLS if c[1] != 'basis']
    rows = [[period, m.scope, weight, m.basis, *[getattr(m, attr) for _, attr, _ in cols[1:]]]
            for (period, weight), metrics in result.metrics.items() for m in metrics
            if m.basis in _SHOWN_BASES and m.scope != '공통' and (m.n_demand or m.n_lt)]
    write_table(sheets['검증비교'], ['구간', '범위', '가중치', '기준', *[h for h, _, _ in cols[1:]]], rows,
                 [None, None, None, None, *[f for _, _, f in cols[1:]]])

    write_table(sheets['적용기준'], ['항목', '내용'], [[k, v] for k, v in meta.items()], [None, None])
    sheets['적용기준'].column_dimensions['A'].width = 24
    sheets['적용기준'].column_dimensions['B'].width = 90

    Path(path).parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(path)
    except PermissionError:
        raise InputError(f'결과 파일을 저장할 수 없습니다(엑셀에서 열려 있으면 닫고 다시 실행): {path}') from None
