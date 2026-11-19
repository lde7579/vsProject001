"""기준 엑셀에서 계산 입력을 한 번에 읽는 공용 로더(report·virtual_data·weight_search가 함께 쓴다)."""
from dataclasses import dataclass

from safety_stock.inventory_days.calc import CalcSettings, SkuResult, build_inputs, compute_sku
from safety_stock.inventory_days.loader import WeeklySales, load_onhand, load_sales, load_summary
from safety_stock.mp_simulation.loader import WeeklyPlan, load_weekly_plan
from safety_stock.purchase_master.grade import GradeResult, compute_grades
from safety_stock.purchase_master.loader import load_purchase_master
from safety_stock.purchase_master.model import PurchaseMasterRow


@dataclass(frozen=True)
class BaseData:
    """기준 엑셀 한 통합문서에서 읽은 값."""
    pm: dict[str, PurchaseMasterRow]
    grades: dict[str, GradeResult]  # 계산 등급(SKU별)
    sales: WeeklySales
    fcst: WeeklyPlan
    onhand: dict[str, float]
    corrections: dict[str, float | None]
    summaries: dict[str, dict]  # '가중평균' | '실적기준' | 'FCST기준' -> load_summary 결과

    @property
    def grade_codes(self) -> dict[str, str | None]:
        """SKU -> 계산 등급 문자(A/B/C, 못 구하면 None)."""
        return {sku: g.grade for sku, g in self.grades.items()}

    def results(self, settings: CalcSettings, pm: dict[str, PurchaseMasterRow] | None = None) -> list[SkuResult]:
        """SKU별 산출 결과. pm을 주면 구매마스터 행만 그것으로 바꿔 계산한다(가중치 탐색용)."""
        inputs = build_inputs(self.pm if pm is None else pm, self.sales, self.fcst, self.onhand,
                              self.corrections, settings, grades=self.grade_codes)
        return [compute_sku(i) for i in inputs]


def load_base(wb) -> BaseData:
    """구매마스터·출고·FCST·현재고·Summary 세 시트를 읽고 등급을 계산한다."""
    pm = load_purchase_master(wb)
    grades = compute_grades(list(pm.values()))
    sales = load_sales(wb)
    fcst = load_weekly_plan(wb, 'SCM FCST')
    onhand = load_onhand(wb)
    weighted = load_summary(wb, 'Summary_가중평균', required=('correction',))
    corrections = {sku: row.get('correction') for sku, row in weighted.items()}
    summaries = {
        '가중평균': weighted,
        '실적기준': load_summary(wb, 'Summary_실적'),
        'FCST기준': load_summary(wb, 'Summary_FCST'),
    }
    return BaseData(pm, grades, sales, fcst, onhand, corrections, summaries)
