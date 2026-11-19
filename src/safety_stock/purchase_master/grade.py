"""재고등급 자동 산출: 백분위 기반 단가·회전율 등급, 점수 합산, 종합 백분위 등급."""
import math
from dataclasses import dataclass

from safety_stock.purchase_master.model import PurchaseMasterRow, normalize_grade

# 백분위 경계: 이 값 이상이면 A, 두 번째 값 이상이면 B, 그 외 C
GRADE_CUTS = (0.75, 0.26)
SCORES = {'price': {'A': 5, 'B': 3, 'C': 1}, 'turnover': {'A': 10, 'B': 5, 'C': 1},
          'lt': {'A': 5, 'B': 3, 'C': 1}, 'quality': {'A': 5, 'B': 3, 'C': 0}}


def percentrank_inc(values: list[float], x: float) -> float:
    """엑셀 PERCENTRANK.INC(소수 셋째 자리 내림). x보다 작은 값의 개수/(n-1), n=1이면 1.0."""
    n = len(values)
    if n <= 1:
        return 1.0
    smaller = sum(v < x for v in values)
    # 부동소수 오차(0.26이 0.2599…로 나오는 경우)를 보정한 뒤 내림
    return math.floor(smaller / (n - 1) * 1000 + 1e-9) / 1000


def grade_from_rank(rank: float) -> str:
    """백분위로 A/B/C를 정한다."""
    if rank >= GRADE_CUTS[0]:
        return 'A'
    return 'B' if rank >= GRADE_CUTS[1] else 'C'


def turnover(fcst, begin, end) -> float | None:
    """회전율 = FCST / ((기초재고 + 기말재고) / 2). 빈 칸은 0으로 본다. 분모가 0이면 None."""
    avg = ((begin or 0) + (end or 0)) / 2
    if avg == 0:
        return None
    return (fcst or 0) / avg


@dataclass(frozen=True)
class GradeResult:
    """한 SKU의 재고등급 계산 결과. 계산 불가 행은 scores·total·total_rank·grade가 None이다."""
    sku: str
    turnover: float | None
    price_rank: float | None
    price_grade: str | None
    turnover_rank: float | None
    turnover_grade: str | None
    lt_grade: str | None
    quality_grade: str | None
    scores: tuple[int, int, int, int] | None
    total: int | None
    total_rank: float | None
    grade: str | None
    memo: str | None


def compute_grades(rows: list[PurchaseMasterRow]) -> dict[str, GradeResult]:
    """모든 SKU 행의 재고등급을 계산한다(입력 순서 유지)."""
    prices = [r.price for r in rows if r.price is not None]
    turns = {r.sku: turnover(r.fcst_qty, r.begin_stock, r.end_stock) for r in rows}
    turn_values = [t for t in turns.values() if t is not None]

    partial = {}
    for r in rows:
        lt, q = normalize_grade(r.lt_grade), normalize_grade(r.quality_grade)
        reasons = [msg for ok, msg in ((r.price is not None, '구매단가없음'), (lt, '리드타임등급없음'),
                                        (q, '품질등급없음')) if not ok]
        price_rank = percentrank_inc(prices, r.price) if r.price is not None else None
        price_grade = grade_from_rank(price_rank) if price_rank is not None else None
        t = turns[r.sku]
        turn_rank = percentrank_inc(turn_values, t) if t is not None else None
        turn_grade = grade_from_rank(turn_rank) if turn_rank is not None else 'C'
        scores = total = None
        if not reasons:
            scores = (SCORES['price'][price_grade], SCORES['turnover'][turn_grade],
                      SCORES['lt'][lt], SCORES['quality'][q])
            total = sum(scores)
        memo = f"등급산출불가({', '.join(reasons)})" if reasons else None
        partial[r.sku] = (t, price_rank, price_grade, turn_rank, turn_grade, lt, q, scores, total, memo)

    totals = [p[8] for p in partial.values() if p[8] is not None]
    out = {}
    for sku, (t, pr, pg, tr, tg, lt, q, scores, total, memo) in partial.items():
        total_rank = percentrank_inc(totals, total) if total is not None else None
        grade = grade_from_rank(total_rank) if total_rank is not None else None
        out[sku] = GradeResult(sku, t, pr, pg, tr, tg, lt, q, scores, total, total_rank, grade, memo)
    return out


def grade_verdict(row: PurchaseMasterRow, g: GradeResult) -> str:
    """엑셀에 적힌 단가·회전율 등급, 종합, 재고등급을 계산값과 대조한다: 일치/불일치/비교불가."""
    pairs = [(normalize_grade(row.excel_price_grade), g.price_grade),
             (normalize_grade(row.excel_turnover_grade), g.turnover_grade),
             (row.excel_total, g.total),
             (normalize_grade(row.grade), g.grade)]
    present = [(e, c) for e, c in pairs if e is not None]
    if not present:
        return '비교불가'
    return '일치' if all(c is not None and e == c for e, c in present) else '불일치'
