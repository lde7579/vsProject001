"""명령줄 진입점: python -m safety_stock BASE.xlsx --actual ACTUAL.xlsx -o OUT.xlsx"""
import argparse
import re
import sys
from pathlib import Path

from safety_stock.common.xlsx import InputError
from safety_stock.inventory_days.calc import CalcSettings
from safety_stock.inventory_days.report import run


def _weeks(text: str) -> range:
    m = re.fullmatch(r'\s*(\d+)\s*-\s*(\d+)\s*', text)
    if not m or int(m.group(1)) > int(m.group(2)):
        raise argparse.ArgumentTypeError(f"주차 범위 형식이 잘못되었습니다: '{text}' (예: 35-43)")
    return range(int(m.group(1)), int(m.group(2)) + 1)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog='safety_stock', description='SKU별 적정 재고일수 산출·대조·검증')
    p.add_argument('base', type=Path, help='요청자 기준 엑셀')
    p.add_argument('--actual', type=Path, required=True, help='실제 출고 엑셀(사후 검증용)')
    p.add_argument('-o', '--out', type=Path, required=True, help='결과 xlsx 경로')
    p.add_argument('--actual-last-week', type=int, default=31, help='실적 마지막 주차(기본 31)')
    p.add_argument('--fcst-weeks', type=int, default=13, help='FCST 사용 주 수(기본 13)')
    p.add_argument('--validation-weeks', type=_weeks, default=range(35, 44), help='검증 주차(기본 35-43)')
    args = p.parse_args(argv)

    settings = CalcSettings(actual_last_week=args.actual_last_week, fcst_weeks=args.fcst_weeks)
    try:
        summary = run(args.base, args.actual, args.out, settings, args.validation_weeks)
    except InputError as e:
        print(f'입력 오류: {e}')
        return 2
    print(f"처리 SKU 수: {summary['sku_count']}")
    print(f"가중평균 엑셀 일치 SKU 수: {summary['match_count_weighted']}")
    print(f"결과 파일: {summary['out']}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
