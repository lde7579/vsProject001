"""명령줄 진입점: python -m safety_stock BASE.xlsx --actual ACTUAL.xlsx -o OUT.xlsx"""
import argparse
import re
import sys
from pathlib import Path

from safety_stock.common.xlsx import InputError
from safety_stock.inventory_days.calc import CalcSettings
from safety_stock.inventory_days.report import run


def _weeks(text: str) -> range | None:
    """'A-B'(A ≤ B)를 range(A, B+1)로. 형식이 틀리면 None."""
    m = re.fullmatch(r'\s*(\d+)\s*-\s*(\d+)\s*', text)
    if not m or int(m.group(1)) > int(m.group(2)):
        return None
    return range(int(m.group(1)), int(m.group(2)) + 1)


def _check_args(args) -> tuple[range | None, str | None]:
    """명령줄 값을 검사해 (검증 주차, 오류 메시지)를 돌려준다. 오류가 없으면 메시지는 None."""
    weeks = _weeks(args.validation_weeks)
    if weeks is None:
        return None, f"--validation-weeks 형식이 잘못되었습니다: '{args.validation_weeks}' (예: 35-43, 앞 주차 ≤ 뒤 주차)"
    if args.fcst_weeks < 1:
        return None, f'--fcst-weeks는 1 이상이어야 합니다: {args.fcst_weeks}'
    if args.actual_last_week >= weeks.start:
        return None, (f'--actual-last-week({args.actual_last_week})는 검증 시작 주차({weeks.start})보다 작아야 합니다'
                      '(검증용 실제 출고가 계산에 섞이지 않도록)')
    return weeks, None


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog='safety_stock', description='SKU별 적정 재고일수 산출·대조·검증')
    p.add_argument('base', type=Path, help='요청자 기준 엑셀')
    p.add_argument('--actual', type=Path, required=True, help='실제 출고 엑셀(사후 검증용)')
    p.add_argument('-o', '--out', type=Path, required=True, help='결과 xlsx 경로')
    p.add_argument('--actual-last-week', type=int, default=31, help='실적 마지막 주차(기본 31)')
    p.add_argument('--fcst-weeks', type=int, default=13, help='FCST 사용 주 수(기본 13)')
    p.add_argument('--validation-weeks', default='35-43', help='검증 주차 A-B(기본 35-43)')
    args = p.parse_args(argv)

    weeks, error = _check_args(args)
    if error:
        print(f'입력 오류: {error}')
        return 2
    settings = CalcSettings(actual_last_week=args.actual_last_week, fcst_weeks=args.fcst_weeks)
    try:
        summary = run(args.base, args.actual, args.out, settings, weeks)
    except InputError as e:
        print(f'입력 오류: {e}')
        return 2
    print(f"처리 SKU 수: {summary['sku_count']}")
    print(f"가중평균 엑셀 일치 SKU 수: {summary['match_count_weighted']}")
    print(f"재고등급 엑셀 일치 SKU 수: {summary['grade_match_count']}/{summary['grade_count']}")
    print(f"결과 파일: {summary['out']}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
