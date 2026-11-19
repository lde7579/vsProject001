"""명령줄: python -m safety_stock.virtual_data SRC.xlsx -o OUT.xlsx"""
import argparse
import sys
from pathlib import Path

from safety_stock.common.xlsx import InputError
from safety_stock.virtual_data.build import build_virtual


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog='safety_stock.virtual_data',
                                description='산식과 어긋난 칸만 채운 가상 보완본 생성(발표용 가상 데이터)')
    p.add_argument('src', type=Path, help='보완본 엑셀(값만 있는 통합문서)')
    p.add_argument('-o', '--out', type=Path, required=True, help='가상 보완본 xlsx 경로')
    args = p.parse_args(argv)
    try:
        changes = build_virtual(args.src, args.out)
    except InputError as e:
        print(f'입력 오류: {e}')
        return 2
    print(f'바꾼 칸 수: {len(changes)}')
    print(f'결과 파일: {args.out}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
