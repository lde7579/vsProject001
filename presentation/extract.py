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
import json
import re
import statistics
import subprocess
import sys
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'out'
DATA = ROOT / 'data'

# 발표 기본 수치: 기준 파일(빈 칸만 산식·다른 시트로 역산해 채운 파일)로 돌린 결과
RESULT = OUT / '결과_산출검증.xlsx'
# 기준 파일(보완내역 시트에 역산해 채운 칸이 기록되어 있다)
BASE_BOOK = DATA / '재고등급_및_안전재고_수준_점검_가중평균수요_적용_20260828_사진전사_보완.xlsx'
# 사후 실제 출고 파일(가중치 탐색 그룹별 기여 계산용)
ACTUAL_BOOK = DATA / 'W44_가중평균수요_반영_상품_발주점검_업로드.xlsx'
# 가중치 탐색 검증 구간(가중치탐색.xlsx 적용기준과 같은 기본값)
HOLDOUT_WEEKS = range(39, 44)
# 가중치 탐색 결과(아직 없을 수 있다)
SEARCH_DEFAULT = OUT / '가중치탐색.xlsx'

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


def grades(path: Path) -> dict:
    """재고등급 시트: 기존 등급과 일치한 수, 등급 분포, 회전율 빈 칸."""
    return grade_summary(sheet_rows(path, '재고등급'))


def grade_summary(rows: list[dict]) -> dict:
    """재고등급 시트 행 요약. match는 4항목(단가·회전율 등급, 종합, 재고등급) 판정 일치 수,
    grade_match는 재고등급만 엑셀과 같은 수(발표의 '기존 등급 일치'는 이 값)."""
    blank = [r for r in rows if r['회전율'] is None]
    return {
        'total': len(rows),
        'match': sum(r['판정'] == '일치' for r in rows),
        'grade_match': sum(r['재고등급'] is not None and r['재고등급'] == r['엑셀재고등급'] for r in rows),
        'verdicts': dict(Counter(r['판정'] for r in rows)),
        'dist': {g: sum(r['재고등급'] == g for r in rows) for g in 'ABC'},
        'not_computable': sum(1 for r in rows if (r['메모'] or '').startswith('등급산출불가')),
        'turnover_blank': len(blank),
        'turnover_blank_grade': dict(Counter(r['회전율등급'] for r in blank)),
        'groups': dict(Counter(r['품목그룹'] for r in rows)),
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


def basis_spread(values: list[float], old: list[float]) -> dict:
    """한 기준의 안전재고일수 분포 요약. shorter는 같은 위치의 기존 값보다 짧은 SKU 수."""
    q = statistics.quantiles(values, n=10, method='inclusive')
    return {
        'median': statistics.median(values), 'min': min(values), 'max': max(values),
        'p10': q[0], 'p90': q[8],
        'distinct': len({round(v, 1) for v in values}),
        'shorter': sum(v < o for v, o in zip(values, old)),
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
    # 네 기준이 모두 있는 SKU로 기준끼리 비교한다
    cols = {'기존수기': '기존안전재고일수', '가중평균': '가중평균_안전재고일수',
            '실적기준': '실적기준_안전재고일수', 'FCST기준': 'FCST기준_안전재고일수'}
    four = [r for r in calc if all(r[c] is not None for c in cols.values())]
    base_old = [r[cols['기존수기']] for r in four]
    bases = {'n': len(four)}
    for name, col in cols.items():
        bases[name] = basis_spread([r[col] for r in four], base_old)
    return {
        'bases': bases,
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
    out = {
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
    for p, b in (('act', '실적기준'), ('fc', 'FCST기준')):
        out[f'{p}_s'] = r[f'{b}_안전재고일수']
        out[f'{p}_u'] = r[f'{b}_적정재고일수']
        out[f'{p}_cover'] = v[f'{b}_LT수요방어']
        out[f'{p}_actual_days'] = v[f'{b}_실제기준재고일수']
    return out


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


def supplement_changes() -> dict | None:
    """기준 파일의 보완내역 시트: 확인되지 않은 빈 칸을 산식·다른 시트 값으로 역산해 채운 칸 수."""
    if not BASE_BOOK.exists():
        return None
    rows = sheet_rows(BASE_BOOK, '보완내역')
    require_columns(rows, ['시트', '셀', '넣은 값', '근거'], '보완내역')
    by_sheet = Counter(r['시트'] for r in rows)
    fcst = [r for r in rows if 'FCST' in str(r['시트'])]
    return {
        'source': rel(BASE_BOOK) + ' › 보완내역',
        'total': len(rows),
        'by_sheet': dict(by_sheet),
        'fcst': len(fcst),
        'fcst_backsolved': sum('역산' in str(r['근거']) for r in fcst),
        'fcst_same_level': sum('같은 수준' in str(r['근거']) for r in fcst),
        'onhand': sum(n for k, n in by_sheet.items() if '현재고' in k),
        'ending': sum(n for k, n in by_sheet.items() if '기말재고' in k),
    }


def test_count() -> int | None:
    """자동 테스트 개수(pytest --collect-only). 세지 못하면 None."""
    try:
        out = subprocess.run([sys.executable, '-m', 'pytest', '--collect-only', '-q'], cwd=ROOT,
                             capture_output=True, text=True, encoding='utf-8', timeout=300).stdout
    except (OSError, subprocess.SubprocessError):
        return None
    m = re.search(r'(\d+) tests? collected', out)
    return int(m.group(1)) if m else None


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

    for p in (RESULT, BASE_BOOK):
        if not p.exists():
            raise SystemExit(f'입력 파일이 없습니다: {p}')

    calc = sheet_rows(RESULT, '산출결과')
    val_rows = sheet_rows(RESULT, '검증_SKU')
    val = {r['SKU']: r for r in val_rows}
    crit = criteria(RESULT)
    g = grades(RESULT)

    data = {
        'generated': date.today().isoformat(),
        'sources': {'result': rel(RESULT), 'base_book': rel(BASE_BOOK), 'search': rel(args.search)},
        'criteria': crit,
        'run_date': str(crit.get('실행 시각', ''))[:10],
        'grades': g,
        'metrics': metrics(RESULT),
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
        'tests': test_count(),
        'weight_search': weight_search(args.search),
    }
    if data['weight_search'] and ACTUAL_BOOK.exists():
        data['weight_search']['contributions'] = weight_contributions(data['weight_search'], calc, val_rows)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding='utf-8')
    ws = '있음' if data['weight_search'] else '없음(산출 예정으로 표시)'
    print(f'{rel(args.out)} 작성: SKU {len(calc)}개, 가중치 탐색 {ws}')


if __name__ == '__main__':
    main()
