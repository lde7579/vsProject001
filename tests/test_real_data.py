"""실제 요청자 엑셀(원본 사진전사본)로 엑셀 값 재현 여부를 확인하는 수용 테스트."""
from pathlib import Path

import openpyxl
import pytest

from safety_stock.inventory_days.calc import CalcSettings, build_inputs, compute_sku
from safety_stock.inventory_days.compare import compare_with_excel, sku_verdicts
from safety_stock.inventory_days.loader import load_onhand, load_sales, load_summary
from safety_stock.inventory_days.validate import VALIDATION_BASES, summarize, validate
from safety_stock.mp_simulation.loader import load_weekly_plan
from safety_stock.purchase_master.loader import load_purchase_master

BASE = Path('data/재고등급_및_안전재고_수준_점검_가중평균수요_적용_20260828_사진전사.xlsx')
ACTUAL = Path('data/W44_가중평균수요_반영_상품_발주점검_업로드.xlsx')
# 요청자 엑셀의 가중평균 시트에서 값이 실적기준 시트와 똑같은 SKU(ETCBID 3 + FLT 52). 엑셀이 이 품목군에 가중평균 대신 실적기준 값을 넣었다.
COPIED_FROM_ACTUAL = {'ETCBID3000MF', 'ETCBID3200WH', 'ETCBID6000MF', 'FLTA15C1A1WH', 'FLTA15C1A3WH', 'FLTA15C2ARWH',
    'FLTA15C2ATWH', 'FLTA16C1A1BL', 'FLTA16C1A3BL', 'FLTA16C1ARZG', 'FLTA16C1ATZG', 'FLTA19P1A4WH', 'FLTA20C1A1WH',
    'FLTA20C1A3WH', 'FLTA20C2ARWH', 'FLTA20C2ATWH', 'FLTA22C1A1BL', 'FLTA22C1A3BL', 'FLTA22C1ARZG', 'FLTA22C1ATZG',
    'FLTA22P1CRBU', 'FLTA22P1CTBU', 'FLTA25C1A1WH', 'FLTA25C1A3WH', 'FLTA25C2ARWH', 'FLTA25C2ATWH', 'FLTACL130PWH',
    'FLTACLP131WH', 'FLTPET131PWH', 'FLTA15C1ATWH', 'FLTA1MDST1GM', 'FLTA1MPTL1GM', 'FLTA1PALG1SG', 'FLTACL120PWH',
    'FLTACL140PWH', 'FLTACL150PWH', 'FLTACL180FWH', 'FLTACL180SWH', 'FLTACL210CBL', 'FLTACL210DWH', 'FLTACL210PWH',
    'FLTACLP122WH', 'FLTACLP160WH', 'FLTACLP190WH', 'FLTACLP201WH', 'FLTACLP251WH', 'FLTACLR11CWH', 'FLTACLR11DWH',
    'FLTACLR11HWH', 'FLTACLV09PWH', 'FLTACLV16PWH', 'FLTACLV32PWH', 'FLTACP110IWH', 'FLTACPP111WH', 'FLTACPP240WH'}
# 원본 전사본에서 FCST 칸이 비어 있는 SKU
FCST_BLANK = {'SENSAWL1KRDW', 'SENSAWL1KRNG', 'SENSAWR1KRDW', 'SENSAWR1KRNG'}


class RealRun:
    """run과 같은 순서로 로드·계산·대조·검증한 결과."""

    def __init__(self):
        wb = openpyxl.load_workbook(BASE, data_only=True)
        actual_wb = openpyxl.load_workbook(ACTUAL, data_only=True)
        weighted = load_summary(wb, 'Summary_가중평균', required=('correction',))
        corrections = {s: r.get('correction') for s, r in weighted.items()}
        summaries = {'가중평균': weighted, '실적기준': load_summary(wb, 'Summary_실적'),
                     'FCST기준': load_summary(wb, 'Summary_FCST')}
        inputs = build_inputs(load_purchase_master(wb), load_sales(wb), load_weekly_plan(wb, 'SCM FCST'),
                              load_onhand(wb), corrections, CalcSettings())
        self.results = [compute_sku(i) for i in inputs]
        self.rows = compare_with_excel(self.results, summaries)
        self.verdicts = sku_verdicts(self.rows)
        self.validation = validate(self.results, load_sales(actual_wb), range(35, 44))
        self.metrics = summarize(self.validation)

    def _demand_rows(self, basis):
        return [r for r in self.rows if r.basis == basis and r.field == 'demand'
                and r.status in ('일치', '불일치')]

    def demand_match(self, basis):
        rows = self._demand_rows(basis)
        return sum(r.status == '일치' for r in rows), len(rows)

    def demand_mismatch_skus(self, basis):
        return {r.sku for r in self._demand_rows(basis) if r.status == '불일치'}


@pytest.fixture(scope='module')
def real_run():
    if not BASE.exists() or not ACTUAL.exists():
        pytest.skip('실데이터 파일이 없습니다')
    return RealRun()


def test_reproduces_excel_weighted(real_run):
    v = real_run.verdicts
    weighted = {s: st for (b, s), st in v.items() if b == '가중평균'}
    assert len(weighted) == 101
    mismatched = {s for s, st in weighted.items() if st == '불일치'}
    # 산식으로 계산된 42개 SKU는 모두 일치, 불일치는 위 두 집합 안에서만
    assert mismatched <= COPIED_FROM_ACTUAL | FCST_BLANK
    assert FCST_BLANK <= mismatched
    assert {s for s in weighted if s not in COPIED_FROM_ACTUAL | FCST_BLANK} == \
        {s for s, st in weighted.items() if st == '일치'} - COPIED_FROM_ACTUAL


def test_actual_and_fcst_basis_demand_match(real_run):
    # 엑셀에 값이 있는 행 기준: 실적기준 수요 59행 모두 일치.
    # FCST기준 46행 중 42행 일치, 불일치 4행은 전사본 FCST 칸이 비어 있는 SEN SKU
    assert real_run.demand_match('실적기준') == (59, 59)
    assert real_run.demand_match('FCST기준') == (42, 46)
    assert real_run.demand_mismatch_skus('FCST기준') == FCST_BLANK


def test_validation_summary_filled(real_run):
    overall = {m.basis: m for m in real_run.metrics if m.scope == '전체'}
    assert set(overall) == set(VALIDATION_BASES)
    # Task 7 결정: 실적·FCST 입력이 없는 SKU는 채점 제외(FCST없음 24, 실적없음+FCST없음 15, 실적파일없음 11) → 126 중 76
    assert overall['가중평균'].n_demand >= 70 and overall['가중평균'].n_lt > 0 and overall['기존수기'].n_lt > 0
