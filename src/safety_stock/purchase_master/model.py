"""구매마스터 영역의 데이터 모델과 재고등급 규칙(읽기 로직은 loader.py)."""
from dataclasses import dataclass

# 재고등급별 서비스 수준 Z값
Z_BY_GRADE = {'A': 1.645, 'B': 1.282, 'C': 1.036}


def normalize_grade(value) -> str | None:
    """등급 값을 'A'/'B'/'C'로 정규화한다. 그 외('#N/A' 포함)는 None."""
    if value is None:
        return None
    g = str(value).strip().upper()
    return g if g in Z_BY_GRADE else None


def z_for_grade(grade: str | None) -> float | None:
    """등급의 Z값. 등급이 없거나 모르는 값이면 None."""
    return Z_BY_GRADE.get(normalize_grade(grade))


@dataclass(frozen=True)
class PurchaseMasterRow:
    """구매마스터 한 SKU의 행. grade는 엑셀에 적힌 재고등급이다(자동 산출 등급이 아님)."""
    sku: str
    group: str | None
    name: str | None
    grade: str | None
    weight_actual: float | None
    weight_fcst: float | None
    lead_time_days: float | None
    current_ss_days: float | None
    # 등급 산출 원재료
    price: float | None = None
    begin_stock: float | None = None
    end_stock: float | None = None
    fcst_qty: float | None = None
    # 입력 등급
    lt_grade: str | None = None
    quality_grade: str | None = None
    # 엑셀에 적힌 값(대조용)
    excel_price_grade: str | None = None
    excel_turnover_grade: str | None = None
    excel_turnover: float | None = None
    excel_total: float | None = None
    # 시트 행 번호
    sheet_row: int | None = None
