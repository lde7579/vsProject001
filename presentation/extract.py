"""발표자료에 쓰는 숫자를 결과 xlsx에서 읽어 presentation/data.json으로 만든다.

숫자는 손으로 적지 않는다. 결과 파일이 바뀌면 이 스크립트와 build.js를 다시 돌린다.

실행(저장소 루트에서):
    uv run python presentation/extract.py
선택 인자:
    --search PATH   가중치 탐색 결과 파일(기본 out/가중치탐색.xlsx, 없으면 null)
    -o PATH         결과 json 경로(기본 presentation/data.json)
"""
from __future__ import annotations

import argparse
import ast
import json
import re
import statistics
import sys
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'out'
DATA = ROOT / 'data'

# 발표 기본 수치: 가상 보완 데이터로 돌린 결과
RESULT_VIRTUAL = OUT / '결과_가상보완.xlsx'
# 가상 보완 전 결과(전후 비교용)
RESULT_BEFORE = OUT / '결과_보완본.xlsx'
# 가상 보완본(바꾼 칸 기록 시트가 들어 있다)
VIRTUAL_BOOK = DATA / '재고등급_및_안전재고_수준_점검_가중평균수요_적용_20260828_가상보완.xlsx'
# 보완본(사진 전사본의 빈 칸을 추정·역산해 채운 파일, 보완내역 시트)
SUPPLEMENT_BOOK = DATA / '재고등급_및_안전재고_수준_점검_가중평균수요_적용_20260828_사진전사_보완.xlsx'
# 사후 실제 출고 파일(가중치 탐색 그룹별 기여 계산용)
ACTUAL_BOOK = DATA / 'W44_가중평균수요_반영_상품_발주점검_업로드.xlsx'
# 가중치 탐색 검증 구간(가중치탐색.xlsx 적용기준과 같은 기본값)
HOLDOUT_WEEKS = range(39, 44)
# 가중치 탐색 결과(아직 없을 수 있다)
SEARCH_DEFAULT = OUT / '가중치탐색.xlsx'
# 실데이터 엑셀 불일치 고정값이 들어 있는 수용 테스트
REAL_TEST = ROOT / 'tests' / 'test_real_data.py'

BASES = ['기존수기', '가중평균', '실적기준', 'FCST기준']
# 재고일수 분포 구간(일). 마지막 구간은 상한 없음
DAY_BINS = [0, 20, 30, 40, 50, 60, 70, 80]
# 안전재고일수 분포 구간(일). 마지막 구간은 상한 없음
SS_BINS = [0, 5, 10, 15, 20, 25, 30, 40, 50]


def rel(path: Path) -> str:
    """저장소 기준 상대 경로(슬래시)."""
    try:
        return path.resolve().relative_to(ROOT).as_posix()
    except ValueError:
        return path.as_posix()


def sheet_rows(path: Path, sheet: str) -> list[dict]:
    """첫 행을 제목으로 보고 시트의 나머지 행을 dict 목록으로 읽는다."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet not in wb.sheetnames:
        raise SystemExit(f'{rel(path)}에 시트 {sheet}가 없습니다')
    it = wb[sheet].iter_rows(values_only=True)
    header = next(it)
    rows = [dict(zip(header, r)) for r in it if any(v is not None for v in r)]
    wb.close()
    return rows


def require_columns(rows: list[dict], cols: list[str], where: str) -> None:
    """예정된 열이 없으면 바로 멈춘다(조용히 빈 값으로 두지 않는다)."""
    if not rows:
        raise SystemExit(f'{where}: 행이 없습니다')
    missing = [c for c in cols if c not in rows[0]]
    if missing:
        raise SystemExit(f'{where}: 열 {missing}이 없습니다. 실제 열: {list(rows[0])}')


def criteria(path: Path) -> dict[str, str]:
    """적용기준 시트(항목, 내용)를 dict로."""
    return {r['항목']: r['내용'] for r in sheet_rows(path, '적용기준')}


def metrics(path: Path) -> dict:
    """검증요약 시트 → {범위: {기준: 지표}}."""
    out: dict = defaultdict(dict)
    for r in sheet_rows(path, '검증요약'):
        out[r['범위']][r['기준']] = {
            'n_demand': r['수요비교SKU수'], 'wape': r['WAPE'], 'bias': r['편향'],
            'n_lt': r['LT검증SKU수'], 'cover': r['결품방어율'], 'days_median': r['실제기준재고일수중앙값'],
        }
    return dict(out)


def compare_counts(path: Path) -> dict:
    """엑셀대조 시트에서 기준별 SKU 판정 개수와 상태별 행 수."""
    rows = sheet_rows(path, '엑셀대조')
    verdict = {}
    for r in rows:
        verdict[(r['기준'], r['SKU'])] = r['SKU판정']
    by_basis = {}
    for b in ['가중평균', '실적기준', 'FCST기준']:
        c = Counter(v for (bb, _), v in verdict.items() if bb == b)
        by_basis[b] = {'total': sum(c.values()), **dict(c)}
    status = Counter(r['상태'] for r in rows)
    no_calc = sorted({r['SKU'] for r in rows if r['상태'] == '계산값없음'})
    return {'by_basis': by_basis, 'status': dict(status), 'no_calc_skus': no_calc}


def real_findings() -> dict:
    """tests/test_real_data.py의 상수와 단언에서 실데이터 불일치 고정값을 읽는다."""
    src = REAL_TEST.read_text(encoding='utf-8')
    sets = {}
    for node in ast.parse(src).body:
        if isinstance(node, ast.Assign) and isinstance(node.value, ast.Set):
            sets[node.targets[0].id] = [ast.literal_eval(e) for e in node.value.elts]

    def find(pattern: str) -> tuple[int, ...]:
        m = re.search(pattern, src)
        if not m:
            raise SystemExit(f'{rel(REAL_TEST)}에서 {pattern}를 찾지 못했습니다')
        return tuple(int(g) for g in m.groups())

    copied = sets['COPIED_FROM_ACTUAL']
    return {
        'source': rel(REAL_TEST),
        'weighted_total': find(r'assert len\(weighted\) == (\d+)')[0],
        'formula_match': find(r'assert len\(formula_skus\) == (\d+)')[0],
        'copied_from_actual': len(copied),
        'copied_etcbid': sum(s.startswith('ETCBID') for s in copied),
        'copied_flt': sum(s.startswith('FLT') for s in copied),
        'fcst_blank': len(sets['FCST_BLANK']),
        'onhand_diff': len(sets['ONHAND_DIFF']),
        'actual_basis_match': find(r"demand_match\('실적기준'\) == \((\d+), (\d+)\)"),
        'fcst_basis_match': find(r"demand_match\('FCST기준'\) == \((\d+), (\d+)\)"),
        'grade_rows': find(r'assert len\(rows\) == (\d+)')[0],
        'turnover_blank': find(r'turnover is None for g in grades.values\(\)\) == (\d+)')[0],
    }


def virtual_changes() -> dict:
    """가상보완_내역 시트에서 바꾼 칸 수와 사유별 개수를 센다.

    바꾼 칸은 구분 = '변경'만 센다. '기록만'은 산식으로 채울 수 없어 사유만 남긴 칸이다.
    """
    rows = sheet_rows(VIRTUAL_BOOK, '가상보완_내역')
    require_columns(rows, ['시트', '셀', '구분', '원래 값', '새 값', '사유'], '가상보완_내역')
    changed = [r for r in rows if r['구분'] == '변경']
    reasons = Counter(r['사유'] for r in rows)
    cat = Counter()
    for r in changed:
        reason = r['사유']
        if reason.startswith('산식 결과로 교체(원래 불일치)'):
            cat['replace_mismatch'] += 1
        elif reason.startswith('산식 결과로 교체'):
            cat['replace_blank'] += 1
        elif reason.startswith('재고등급'):
            cat['grade_recalc'] += 1
        elif reason.startswith('회전율 분모 0'):
            cat['virtual_stock'] += 1
        elif reason.startswith('가상 재고로 회전율 순위'):
            cat['rank_population'] += 1
        else:
            cat['other'] += 1
    stock_rows = {re.sub(r'[A-Z]+', '', r['셀']) for r in changed if r['사유'].startswith('회전율 분모 0')}
    rank_rows = {re.sub(r'[A-Z]+', '', r['셀']) for r in changed if r['사유'].startswith('가상 재고로 회전율 순위')}
    median_turn = re.search(r'중앙회전율 ([\d.]+)', next((k for k in reasons if k.startswith('회전율 분모 0')), ''))
    return {
        'source': rel(VIRTUAL_BOOK) + ' › 가상보완_내역',
        'total': len(changed),
        'recorded_only': sum(r['구분'] == '기록만' for r in rows),
        'replace': cat['replace_mismatch'] + cat['replace_blank'],
        'replace_mismatch': cat['replace_mismatch'],
        'replace_blank': cat['replace_blank'],
        'grade_recalc': cat['grade_recalc'],
        'virtual_stock': cat['virtual_stock'],
        'virtual_stock_rows': len(stock_rows),
        'rank_population': cat['rank_population'],
        'rank_population_rows': len(rank_rows),
        'median_turnover': float(median_turn.group(1)) if median_turn else None,
        'other': cat['other'],
        'by_sheet': dict(Counter(r['시트'] for r in changed)),
        'reasons': dict(reasons),
    }


def grades(path_after: Path, path_before: Path) -> dict:
    """재고등급 시트: 엑셀 일치 수, 등급 분포, 회전율 빈 칸(가상 보완 전)."""
    after = sheet_rows(path_after, '재고등급')
    before = sheet_rows(path_before, '재고등급')
    blank_before = [r for r in before if r['회전율'] is None]
    blank_skus = {r['SKU'] for r in blank_before}
    b_by = {r['SKU']: r for r in before}
    # 가상 재고를 넣지 않은 SKU인데도 회전율 순위 모집단이 바뀌어 등급이 달라진 SKU
    others = [r for r in after if r['SKU'] not in blank_skus and r['SKU'] in b_by]
    turn_changed = sorted(r['SKU'] for r in others if r['회전율등급'] != b_by[r['SKU']]['회전율등급'])
    grade_changed = sorted(r['SKU'] for r in others if r['재고등급'] != b_by[r['SKU']]['재고등급'])
    return {
        'unfilled_turnover_grade_changed': turn_changed,
        'unfilled_grade_changed': grade_changed,
        'unfilled_grade_changes': [f"{s} {b_by[s]['재고등급']}→{next(r['재고등급'] for r in after if r['SKU'] == s)}"
                                   for s in grade_changed],
        'total': len(after),
        'match': sum(r['판정'] == '일치' for r in after),
        'verdicts': dict(Counter(r['판정'] for r in after)),
        'dist': {g: sum(r['재고등급'] == g for r in after) for g in 'ABC'},
        'not_computable': sum(1 for r in after if (r['메모'] or '').startswith('등급산출불가')),
        'turnover_blank_before': len(blank_before),
        'turnover_blank_before_grade': dict(Counter(r['회전율등급'] for r in blank_before)),
        'turnover_blank_after': sum(r['회전율'] is None for r in after),
        'match_before': sum(r['판정'] == '일치' for r in before),
        'total_before': len(before),
        'groups': dict(Counter(r['품목그룹'] for r in after)),
    }


def bins(values: list[float], edges: list[int] = DAY_BINS) -> list[int]:
    """구간별 개수(edges의 마지막 구간은 상한 없음)."""
    counts = [0] * len(edges)
    for v in values:
        i = max(k for k, lo in enumerate(edges) if v >= lo)
        counts[i] += 1
    return counts


def bin_labels(edges: list[int] = DAY_BINS) -> list[str]:
    labels = []
    for k, lo in enumerate(edges):
        if k == 0:
            labels.append(f'{edges[1]}일 미만')
        elif k == len(edges) - 1:
            labels.append(f'{lo}일 이상')
        else:
            labels.append(f'{lo}~{edges[k + 1]}')
    return labels


def days_distribution(calc: list[dict]) -> dict:
    """기존 수기 적정재고일수와 가중평균 적정재고일수 분포(둘 다 있는 SKU)."""
    both = [r for r in calc if r['기존수기_적정재고일수'] is not None and r['가중평균_적정재고일수'] is not None]
    old = [r['기존수기_적정재고일수'] for r in both]
    new = [r['가중평균_적정재고일수'] for r in both]
    old_ss = Counter(r['기존안전재고일수'] for r in calc if r['기존안전재고일수'] is not None)
    new_ss = [r['가중평균_안전재고일수'] for r in both if r['가중평균_안전재고일수'] is not None]
    return {
        'n': len(both),
        'labels': bin_labels(),
        'old': bins(old), 'new': bins(new),
        'old_median': statistics.median(old), 'new_median': statistics.median(new),
        'new_min': min(new), 'new_max': max(new),
        'shorter': sum(n < o for o, n in zip(old, new)),
        'old_distinct': len(set(old)), 'new_distinct': len({round(v, 1) for v in new}),
        'old_safety_levels': {str(k): v for k, v in sorted(old_ss.items())},
        'old_safety_n': sum(old_ss.values()),
        'new_safety_median': statistics.median(new_ss) if new_ss else None,
        'new_safety_min': min(new_ss) if new_ss else None,
        'new_safety_max': max(new_ss) if new_ss else None,
    }


def safety_distribution(calc: list[dict]) -> dict:
    """기존 안전재고일수(MP 일별 안전재고)와 산식 안전재고일수 분포.

    MP에 올릴 값은 적정 재고일수가 아니라 안전재고일수(안전재고 ÷ 일평균수요)다.
    비교는 두 값이 모두 있는 SKU로 한다.
    """
    both = [r for r in calc if r['기존안전재고일수'] is not None and r['가중평균_안전재고일수'] is not None]
    old = [r['기존안전재고일수'] for r in both]
    new = [r['가중평균_안전재고일수'] for r in both]
    all_new = [r['가중평균_안전재고일수'] for r in calc if r['가중평균_안전재고일수'] is not None]
    old_all = [r['기존안전재고일수'] for r in calc if r['기존안전재고일수'] is not None]
    return {
        'n': len(both), 'labels': bin_labels(SS_BINS), 'old': bins(old, SS_BINS), 'new': bins(new, SS_BINS),
        'old_median': statistics.median(old), 'new_median': statistics.median(new),
        'new_min': min(new), 'new_max': max(new), 'shorter': sum(n < o for o, n in zip(old, new)),
        'old_levels_n': len(set(old_all)), 'old_all_n': len(old_all),
        'all_new_n': len(all_new), 'all_new_min': min(all_new), 'all_new_max': max(all_new),
    }


def representative_case(calc: list[dict], val: dict[str, dict]) -> dict:
    """대표 사례: 메모가 없고 두 기준 모두 결품 방어가 채점된 SKU 중,
    기존 안전재고일수가 가장 긴 SKU들 가운데 실적평균이 가장 큰 SKU."""
    cand = [r for r in calc if not r['점검메모'] and r['기존안전재고일수'] is not None
            and val[r['SKU']]['가중평균_LT수요방어'] and val[r['SKU']]['기존수기_LT수요방어']]
    longest = max(r['기존안전재고일수'] for r in cand)
    r = max((r for r in cand if r['기존안전재고일수'] == longest), key=lambda r: r['실적평균'])
    rule = '점검메모가 없고 결품 방어가 채점된 SKU 중 기존 안전재고일수가 가장 긴 SKU, 그중 실적평균이 가장 큰 SKU'
    return case_dict(r, val[r['SKU']], rule)


def counter_case(calc: list[dict], val: dict[str, dict]) -> dict:
    """반대 사례: 메모가 없고 산식(가중평균)이 결품을 막지 못한 SKU 중 오차기여율이 가장 큰 SKU."""
    cand = [r for r in calc if not r['점검메모'] and r['기존안전재고일수'] is not None
            and val[r['SKU']]['가중평균_LT수요방어'] == '결품' and val[r['SKU']]['기존수기_LT수요방어']
            and val[r['SKU']]['가중평균_오차기여율'] is not None]
    r = max(cand, key=lambda r: val[r['SKU']]['가중평균_오차기여율'])
    rule = '점검메모가 없고 산식이 리드타임 수요를 덮지 못한(방어 실패) SKU 중 가중평균 오차기여율이 가장 큰 SKU'
    return case_dict(r, val[r['SKU']], rule)


def case_dict(r: dict, v: dict, rule: str) -> dict:
    """사례 장에 쓰는 값(산출결과 한 행 + 검증_SKU 한 행)."""
    return {
        'rule': rule,
        'sku': r['SKU'], 'group': r['품목그룹'], 'name': r['품목명'], 'grade': r['재고등급'],
        'wh': r['실적가중치'], 'wf': r['FCST가중치'], 'lt': r['리드타임(일)'], 'onhand': r['현재고'],
        'hm': r['실적평균'], 'fm': r['FCST평균'], 'g': r['가중평균_수요'], 'sigma': r['가중평균_σ'],
        'k': r['가중평균_LT평균수요'], 'l': r['가중평균_LT수요편차'], 'z': r['가중평균_Z'],
        't': r['가중평균_안전재고'], 'rop': r['가중평균_ROP'], 'u': r['가중평균_적정재고일수'],
        's': r['가중평균_안전재고일수'], 'e': r['가중평균_현재고일수'], 'order': r['가중평균_발주'],
        'old_ss_days': r['기존안전재고일수'], 'old_days': r['기존수기_적정재고일수'],
        'actual_week': v['실제주평균'], 'actual_lt': v['실제LT수요'],
        'old_target': v['기존수기_목표재고량'], 'new_target': v['가중평균_목표재고량'],
        'old_cover': v['기존수기_LT수요방어'], 'new_cover': v['가중평균_LT수요방어'],
        'old_actual_days': v['기존수기_실제기준재고일수'], 'new_actual_days': v['가중평균_실제기준재고일수'],
        'error': v['가중평균_수요오차'],
    }


def validation_counts(val: list[dict], calc: list[dict]) -> dict:
    """검증_SKU: 전체 SKU, 실적파일없음, 기준별 채점 제외 사유와 LT 채점 수."""
    lt = {r['SKU']: r['리드타임(일)'] for r in calc}
    out = {
        'total': len(val),
        'memo': dict(Counter(r['메모'] or '' for r in val)),
        'lt_over_window': sum(1 for r in val if r['메모'] is None and (lt.get(r['SKU']) or 0) > 63),
        'by_basis': {},
    }
    for b in BASES:
        out['by_basis'][b] = {
            'excluded': dict(Counter(r[f'{b}_제외사유'] for r in val if r[f'{b}_제외사유'])),
            'covered': sum(r[f'{b}_LT수요방어'] == '방어' for r in val),
            'short': sum(r[f'{b}_LT수요방어'] == '결품' for r in val),
            'demand_scored': sum(r[f'{b}_수요오차'] is not None for r in val),
        }
    return out


def top_contributors(val: list[dict], n: int = 5) -> dict:
    """가중평균 오차기여율 상위 SKU."""
    rows = sorted((r for r in val if r['가중평균_오차기여율'] is not None), key=lambda r: -r['가중평균_오차기여율'])
    top = [{'sku': r['SKU'], 'group': r['품목그룹'], 'forecast': r['가중평균_예측주수요'],
            'actual': r['실제주평균'], 'error': r['가중평균_수요오차'], 'share': r['가중평균_오차기여율']}
           for r in rows[:n]]
    by_group: dict[str, float] = defaultdict(float)
    for r in rows:
        by_group[r['품목그룹'] or '미분류'] += r['가중평균_오차기여율']
    return {'rows': top, 'share_sum': sum(t['share'] for t in top), 'n_scored': len(rows),
            'group_share': dict(by_group)}


def group_sums(calc: list[dict], val: dict[str, dict]) -> dict:
    """공통 범위(세 산식 기준 모두 수요오차가 있는 SKU)의 그룹별 주수요 합계."""
    out = {}
    for r in calc:
        v = val[r['SKU']]
        if not all(v[f'{b}_수요오차'] is not None for b in ['가중평균', '실적기준', 'FCST기준']):
            continue
        g = out.setdefault(r['품목그룹'], {'n': 0, 'hist': 0.0, 'fcst': 0.0, 'wavg': 0.0, 'actual': 0.0})
        g['n'] += 1
        g['hist'] += r['실적평균'] or 0
        g['fcst'] += r['FCST평균'] or 0
        g['wavg'] += r['가중평균_수요']
        g['actual'] += v['실제주평균']
    return out


def group_weights(calc: list[dict], grade_groups: dict[str, int]) -> list[dict]:
    """그룹별 현재 실적·FCST 가중치(최빈값)와 구매마스터 SKU 수."""
    w = defaultdict(Counter)
    for r in calc:
        if r['품목그룹'] and r['실적가중치'] is not None:
            w[r['품목그룹']][(r['실적가중치'], r['FCST가중치'])] += 1
    out = []
    for g, n in sorted(grade_groups.items(), key=lambda kv: (-kv[1], kv[0])):
        (wh, wf), _ = sorted(w[g].items(), key=lambda kv: (-kv[1], kv[0]))[0]
        out.append({'group': g, 'skus': n, 'wh': wh, 'wf': wf})
    return out


def memos(calc: list[dict]) -> dict:
    """산출결과 점검메모별 SKU 수."""
    c = Counter(t.strip() for r in calc if r['점검메모'] for t in r['점검메모'].split(','))
    return {'counts': dict(c), 'with_memo': sum(1 for r in calc if r['점검메모']), 'total': len(calc),
            'order_needed': sum(r['가중평균_발주'] == '발주필요' for r in calc),
            'order_ok': sum(r['가중평균_발주'] == '정상' for r in calc)}


def supplement_changes() -> dict:
    """보완본의 보완내역 시트: 사진 전사본의 빈 칸을 추정·역산으로 채운 칸 수."""
    if not SUPPLEMENT_BOOK.exists():
        return None
    rows = sheet_rows(SUPPLEMENT_BOOK, '보완내역')
    require_columns(rows, ['시트', '셀', '넣은 값', '근거'], '보완내역')
    by_sheet = Counter(r['시트'] for r in rows)
    fcst = [r for r in rows if 'FCST' in str(r['시트'])]
    return {
        'source': rel(SUPPLEMENT_BOOK) + ' › 보완내역',
        'total': len(rows),
        'by_sheet': dict(by_sheet),
        'fcst': len(fcst),
        'fcst_backsolved': sum('역산' in str(r['근거']) for r in fcst),
        'fcst_same_level': sum('같은 수준' in str(r['근거']) for r in fcst),
        'onhand': sum(n for k, n in by_sheet.items() if '현재고' in k),
        'ending': sum(n for k, n in by_sheet.items() if '기말재고' in k),
    }


def holdout_actuals() -> dict[str, float]:
    """검증 구간(W39~43) SKU별 실제 주평균. validate와 같게 빈 주차는 0으로 보고 주 수로 나눈다."""
    sys.path.insert(0, str(ROOT / 'src'))
    from safety_stock.inventory_days.loader import load_sales  # 프로젝트 로더를 그대로 쓴다

    sales = load_sales(openpyxl.load_workbook(ACTUAL_BOOK, read_only=True, data_only=True))
    cols = [sales.week_numbers.index(w) for w in HOLDOUT_WEEKS]
    out = {}
    for sku, series in sales.by_sku.items():
        out[sku] = statistics.fmean([(series[i] if i < len(series) and series[i] is not None else 0.0) for i in cols])
    return out


def weight_contributions(ws: dict, calc: list[dict], val_rows: list[dict]) -> dict | None:
    """검증 구간 Σ|오차| 감소분의 그룹별 몫(SKU 단위로 직접 계산).

    예측 주수요 = 실적평균 × w + FCST평균 × (1 − w). 현재 w는 구매마스터 값, 탐색 w는 선택가중치 시트 값이다.
    대상은 가중평균 수요오차가 채점된 SKU(검증요약 '전체' 범위와 같은 SKU)다.
    다시 계산한 전체 WAPE가 가중치탐색.xlsx 값과 맞는지도 함께 남긴다.
    """
    chosen = {c['group']: c['chosen'] for c in ws['choices']}
    actual = holdout_actuals()
    by = {r['SKU']: r for r in calc}
    cur_err: dict[str, float] = defaultdict(float)
    new_err: dict[str, float] = defaultdict(float)
    act_sum = 0.0
    n = 0
    for v in val_rows:
        if v['가중평균_수요오차'] is None or v['SKU'] not in actual:
            continue
        r = by[v['SKU']]
        g = r['품목그룹']
        hm, fm = (r['실적평균'] or 0.0), (r['FCST평균'] or 0.0)
        a = actual[v['SKU']]
        w_now = r['실적가중치']
        w_new = chosen.get(g, w_now)
        cur_err[g] += abs(hm * w_now + fm * (1 - w_now) - a)
        new_err[g] += abs(hm * w_new + fm * (1 - w_new) - a)
        act_sum += a
        n += 1
    drop = sum(cur_err.values()) - sum(new_err.values())
    groups = {g: {'abs_err_drop': cur_err[g] - new_err[g], 'share': (cur_err[g] - new_err[g]) / drop if drop else None}
              for g in cur_err}
    return {
        'n': n,
        'wape_current_recomputed': sum(cur_err.values()) / act_sum if act_sum else None,
        'wape_search_recomputed': sum(new_err.values()) / act_sum if act_sum else None,
        'total_abs_err_drop': drop,
        'groups': groups,
        'note': 'SKU별 예측 = 실적평균×w + FCST평균×(1−w), 검증 구간 실제 주평균과의 |오차| 합을 그룹별로 비교',
    }


def weight_search(path: Path) -> dict | None:
    """가중치 탐색 결과(없으면 None). 시트·열은 구현 계획의 report.py 설명을 따른다."""
    if not path.exists():
        return None
    cand = sheet_rows(path, '후보성능')
    require_columns(cand, ['범위', '실적가중치', 'SKU수', 'WAPE'], '후보성능')
    choice = sheet_rows(path, '선택가중치')
    require_columns(choice, ['품목그룹', 'SKU수', '현재가중치', '선택가중치', '선택방식',
                             '현재WAPE(선택구간)', '선택WAPE(선택구간)'], '선택가중치')
    comp = sheet_rows(path, '검증비교')
    require_columns(comp, ['구간', '범위', '가중치', '기준', '수요비교SKU수', 'WAPE', '편향',
                           'LT검증SKU수', '결품방어율', '실제기준재고일수중앙값'], '검증비교')
    return {
        'source': rel(path),
        'candidates': [{'scope': r['범위'], 'w': r['실적가중치'], 'n': r['SKU수'], 'wape': r['WAPE']} for r in cand],
        'choices': [{'group': r['품목그룹'], 'n': r['SKU수'], 'current': r['현재가중치'], 'chosen': r['선택가중치'],
                     'method': r['선택방식'], 'wape_current': r['현재WAPE(선택구간)'],
                     'wape_chosen': r['선택WAPE(선택구간)']} for r in choice],
        'compare': [{'period': r['구간'], 'scope': r['범위'], 'weights': r['가중치'], 'basis': r['기준'],
                     'n_demand': r['수요비교SKU수'], 'wape': r['WAPE'], 'bias': r['편향'], 'n_lt': r['LT검증SKU수'],
                     'cover': r['결품방어율'], 'days_median': r['실제기준재고일수중앙값']} for r in comp],
        'criteria': criteria(path),
    }


def main() -> None:
    ap = argparse.ArgumentParser(description='발표 숫자 추출')
    ap.add_argument('--search', type=Path, default=SEARCH_DEFAULT)
    ap.add_argument('-o', '--out', type=Path, default=ROOT / 'presentation' / 'data.json')
    args = ap.parse_args()

    for p in (RESULT_VIRTUAL, RESULT_BEFORE, VIRTUAL_BOOK, REAL_TEST):
        if not p.exists():
            raise SystemExit(f'입력 파일이 없습니다: {p}')

    calc = sheet_rows(RESULT_VIRTUAL, '산출결과')
    val_rows = sheet_rows(RESULT_VIRTUAL, '검증_SKU')
    val = {r['SKU']: r for r in val_rows}
    crit = criteria(RESULT_VIRTUAL)
    m_after, m_before = metrics(RESULT_VIRTUAL), metrics(RESULT_BEFORE)
    g = grades(RESULT_VIRTUAL, RESULT_BEFORE)

    data = {
        'generated': date.today().isoformat(),
        'sources': {
            'virtual': rel(RESULT_VIRTUAL), 'before': rel(RESULT_BEFORE),
            'virtual_book': rel(VIRTUAL_BOOK), 'search': rel(args.search), 'real_test': rel(REAL_TEST),
        },
        'criteria': crit,
        'run_date': str(crit.get('실행 시각', ''))[:10],
        'real': real_findings(),
        'virtual_changes': virtual_changes(),
        'compare': {'before': compare_counts(RESULT_BEFORE), 'after': compare_counts(RESULT_VIRTUAL)},
        'grades': g,
        'metrics': m_after,
        'metrics_same_as_before': m_after == m_before,
        # 가상 보완 전후로 값이 달라진 검증요약 칸 수(범위·기준·지표 단위)
        'metrics_changed_cells': sum(1 for sc in m_after for b in m_after[sc] for k in m_after[sc][b]
                                     if m_before.get(sc, {}).get(b, {}).get(k) != m_after[sc][b][k]),
        'days': days_distribution(calc),
        'case': representative_case(calc, val),
        'counter_case': counter_case(calc, val),
        'safety': safety_distribution(calc),
        'supplement': supplement_changes(),
        'validation': validation_counts(val_rows, calc),
        'top': top_contributors(val_rows),
        'group_sums': group_sums(calc, val),
        'weights': group_weights(calc, g['groups']),
        'memos': memos(calc),
        'sku_total': len(calc),
        'weight_search': weight_search(args.search),
    }
    # 실적기준 값과 같은 SKU의 품목그룹(코드 앞글자와 그룹이 다를 수 있다: ETCBID 코드는 FLT 그룹)
    copied = ast.literal_eval(re.search(r'COPIED_FROM_ACTUAL = (\{.*?\})', REAL_TEST.read_text(encoding='utf-8'), re.S).group(1))
    grp = {r['SKU']: r['품목그룹'] for r in calc}
    data['real']['copied_groups'] = dict(Counter(grp.get(s) or '미분류' for s in copied))
    if data['weight_search'] and ACTUAL_BOOK.exists():
        data['weight_search']['contributions'] = weight_contributions(data['weight_search'], calc, val_rows)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding='utf-8')
    ws = '있음' if data['weight_search'] else '없음(산출 예정으로 표시)'
    print(f'{rel(args.out)} 작성: SKU {len(calc)}개, 가중치 탐색 {ws}')


if __name__ == '__main__':
    main()
