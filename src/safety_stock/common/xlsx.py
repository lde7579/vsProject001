"""요청자 엑셀을 읽을 때 공통으로 쓰는 유틸."""
import math
import re


class InputError(Exception):
    """사용자에게 그대로 보여줄 입력 오류(시트·열 이름 포함 메시지)."""


def norm_sku(v) -> str | None:
    """SKU를 문자열화하고 앞뒤 공백 제거 후 대문자로 만든다. 빈 값은 None."""
    if v is None:
        return None
    s = str(v).strip().upper()
    return s or None


def norm_header(v) -> str:
    """제목 셀의 공백·줄바꿈을 모두 제거한다."""
    if v is None:
        return ''
    return re.sub(r'\s+', '', str(v))


def to_number(v) -> float | None:
    """표시용 값을 숫자로 바꾼다. 숫자로 볼 수 없거나 유한하지 않으면(nan·inf) None."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) if math.isfinite(v) else None
    if not isinstance(v, str):
        return None
    s = v.strip().replace(',', '')
    if not s or s.startswith('#'):
        return None
    negative = False
    if s.startswith('(') and s.endswith(')'):
        negative = True
        s = s[1:-1].strip()
    percent = s.endswith('%')
    if percent:
        s = s[:-1].strip()
    try:
        n = float(s)
    except ValueError:
        return None
    if not math.isfinite(n):
        return None
    if percent:
        n /= 100
    return -n if negative else n


def find_sheet(wb, *keywords: str):
    """정확히 같은 이름의 시트를 우선, 없으면 키워드가 이름에 포함된 첫 시트를 찾는다."""
    for kw in keywords:
        if kw in wb.sheetnames:
            return wb[kw]
    for kw in keywords:
        for name in wb.sheetnames:
            if kw in name:
                return wb[name]
    raise InputError(
        f"시트를 찾을 수 없습니다(찾은 키워드: {', '.join(keywords)}). "
        f"있는 시트: {', '.join(wb.sheetnames)}"
    )


def find_header_row(ws, required: str, max_scan: int = 30) -> int:
    """required 제목 셀이 있는 첫 행 번호를 위에서부터 찾는다."""
    target = norm_header(required)
    for row in ws.iter_rows(min_row=1, max_row=max_scan):
        for cell in row:
            if norm_header(cell.value) == target:
                return cell.row
    raise InputError(f"시트 '{ws.title}'에서 제목 행('{required}')을 찾을 수 없습니다")


def header_map(ws, row: int) -> dict[str, int]:
    """정규화한 제목 -> 열 번호(1부터). 같은 제목은 처음 나온 열."""
    result: dict[str, int] = {}
    for cell in ws[row]:
        key = norm_header(cell.value)
        if key and key not in result:
            result[key] = cell.column
    return result


def header_cols(ws, row: int) -> dict[str, list[int]]:
    """정규화한 제목 -> 그 제목이 나오는 모든 열 번호(왼쪽부터)."""
    result: dict[str, list[int]] = {}
    for cell in ws[row]:
        key = norm_header(cell.value)
        if key:
            result.setdefault(key, []).append(cell.column)
    return result


def require_cols(hmap: dict[str, int], names: list[str], sheet: str) -> list[int]:
    """names에 해당하는 열 번호를 순서대로 돌려준다. 없는 열이 있으면 InputError."""
    missing = [n for n in names if norm_header(n) not in hmap]
    if missing:
        raise InputError(f"시트 '{sheet}'에 열이 없습니다: {', '.join(missing)}")
    return [hmap[norm_header(n)] for n in names]
