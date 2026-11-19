"""명령줄 진입점: python -m safety_stock.weight_search BASE.xlsx --actual ACTUAL.xlsx -o OUT.xlsx"""
import argparse
import sys
from datetime import datetime
from pathlib import Path

from safety_stock.common.xlsx import InputError, parse_weeks
from safety_stock.inventory_days.calc import CalcSettings
from safety_stock.weight_search.report import write_search_report
from safety_stock.weight_search.search import CANDIDATES, MIN_GROUP_SKUS, SearchResult, run_search


def _label(weeks: range) -> str:
    return f'WK{weeks.start:02d}~WK{weeks.stop - 1:02d}'


def _meta(base: Path, actual: Path, select: range, holdout: range, full: range) -> dict[str, str]:
    return {
        '실행 시각': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        '기준 파일': base.name,
        '실제 출고 파일': actual.name,
        '선택 구간': f'{_label(select)} 실제 출고로 가중치를 고른다',
        '검증 구간': f'{_label(holdout)} 실제 출고로 현재와 탐색 가중치를 비교한다(선택에 쓰지 않음)',
        '참고 전체 구간': f'{_label(full)} (선택 구간이 섞여 있어 참고용)',
        '후보': f'실적가중치 {CANDIDATES[0]}~{CANDIDATES[-1]} (0.1 간격), FCST가중치 = 1 - 실적가중치',
        '단위': '품목그룹. 그룹 안의 모든 SKU에 같은 가중치',
        '고르는 기준': '선택 구간 가중평균 수요 WAPE 최소',
        '동점 규칙': '차이 1e-12 이하면 현재 가중치에 가까운 쪽, 그래도 같으면 작은 쪽',
        '작은 그룹 규칙': f'선택 구간에서 수요오차가 있는 SKU가 {MIN_GROUP_SKUS}개 미만이면 모든 SKU를 합친 WAPE 최소 가중치(공통최적)',
        '현재유지': '그룹 또는 전체의 WAPE를 구할 수 없으면(실제 합 0 등) 현재 가중치를 유지',
        '재고등급': '계산 등급 사용(2단계)',
        '검증비교 범위': "검증비교의 그룹 범위는 1단계 검증요약의 '공통' 범위(세 산식 기준 모두 오차가 있는 SKU)를 그룹으로 거른 것이라, 선택가중치의 SKU 수와 다를 수 있다",
        '한계': '선택 구간 4주뿐이고 통계적 유의성 검정이 없다. 개선 주장은 검증 구간 수치로만 한다',
    }


def _pct(metrics, basis: str = '가중평균') -> float | None:
    return next((m.wape for m in metrics if m.scope == '전체' and m.basis == basis), None)


def _fmt(v: float | None) -> str:
    return '-' if v is None else f'{v * 100:.1f}%'


def _print(result: SearchResult) -> None:
    for c in result.choices:
        print(f'{c.group}: {c.current} → {c.chosen} ({c.method})')
    cur, new = (_pct(result.metrics[('검증', k)]) for k in ('현재', '탐색'))
    print(f'검증 구간 가중평균 WAPE: 현재 {_fmt(cur)} → 탐색 {_fmt(new)}')


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog='safety_stock.weight_search', description='품목그룹별 실적가중치 탐색')
    p.add_argument('base', type=Path, help='기준 엑셀(발표용은 가상 보완본)')
    p.add_argument('--actual', type=Path, required=True, help='실제 출고 엑셀')
    p.add_argument('-o', '--out', type=Path, required=True, help='결과 xlsx 경로')
    p.add_argument('--select-weeks', default='35-38', help='선택 구간 A-B(기본 35-38)')
    p.add_argument('--holdout-weeks', default='39-43', help='검증 구간 A-B(기본 39-43)')
    args = p.parse_args(argv)

    select, holdout = parse_weeks(args.select_weeks), parse_weeks(args.holdout_weeks)
    error = None
    if select is None:
        error = f"--select-weeks 형식이 잘못되었습니다: '{args.select_weeks}' (예: 35-38, 앞 주차 ≤ 뒤 주차)"
    elif holdout is None:
        error = f"--holdout-weeks 형식이 잘못되었습니다: '{args.holdout_weeks}' (예: 39-43, 앞 주차 ≤ 뒤 주차)"
    elif holdout.start < select.stop:
        error = f'--holdout-weeks({args.holdout_weeks})는 --select-weeks({args.select_weeks}) 뒤여야 합니다(겹치면 선택 데이터가 검증에 섞임)'
    elif select.start <= CalcSettings().actual_last_week:
        error = (f'--select-weeks 시작({select.start})은 실적 마지막 주차({CalcSettings().actual_last_week})보다 '
                 '뒤여야 합니다(선택 구간 실제 출고가 계산 입력에 섞임)')
    if error:
        print(f'입력 오류: {error}')
        return 2

    full = range(select.start, holdout.stop)
    try:
        for src in (args.base, args.actual):
            if args.out.resolve() == src.resolve():
                raise InputError(f'결과 파일 경로가 입력 파일과 같습니다: {args.out}')
        result = run_search(args.base, args.actual, CalcSettings(), select, holdout, full)
        write_search_report(args.out, result, _meta(args.base, args.actual, select, holdout, full))
    except InputError as e:
        print(f'입력 오류: {e}')
        return 2
    _print(result)
    print(f'결과 파일: {args.out}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
