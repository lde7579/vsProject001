"""실제 요청자 엑셀(원본 사진전사본)로 엑셀 값 재현 여부를 확인하는 수용 테스트."""
from pathlib import Path

import openpyxl
import pytest

from safety_stock.inventory_days.calc import CalcSettings, build_inputs, compute_sku
from safety_stock.inventory_days.compare import compare_with_excel, sku_verdicts
from safety_stock.inventory_days.base import load_base
from safety_stock.inventory_days.loader import load_sales
from safety_stock.inventory_days.validate import VALIDATION_BASES, summarize, validate
from safety_stock.purchase_master.grade import compute_grades, grade_verdict
from safety_stock.virtual_data.build import build_virtual

DATA = Path(__file__).resolve().parents[1] / 'data'
BASE = DATA / Path('재고등급_및_안전재고_수준_점검_가중평균수요_적용_20260828_사진전사.xlsx')
SUPPLEMENT = DATA / '재고등급_및_안전재고_수준_점검_가중평균수요_적용_20260828_사진전사_보완.xlsx'
VIRTUAL = DATA / '재고등급_및_안전재고_수준_점검_가중평균수요_적용_20260828_가상보완.xlsx'
ACTUAL = DATA / Path('W44_가중평균수요_반영_상품_발주점검_업로드.xlsx')
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
# 가중평균 Summary의 현재고가 계산 쪽 현재고(현재고 시트 합)와 다른 SKU: 값 차이 21 + 현재고 시트에 없는 3
# (FLTACL140PWH, MATQM730RZOM, MATQH530RKBE). 최종 리뷰에서 현재고를 대조 필드에 넣으며 드러났다.
# 리뷰 메모는 25건으로 적었으나 실제 대조 결과는 24건(세 기준 합집합도 24건)이어서 실제 결과를 고정한다.
ONHAND_DIFF = {'ACL131T0CJDG', 'ETCBID3200WH', 'ETCDRAINHOS1', 'ETCDRAINTANK', 'FLTACL140PWH', 'MATKH530RKBE',
    'MATKM230RSBR', 'MATKM430RLWH', 'MATQH510RKWH', 'MATQH520RKIV', 'MATQH530RKBE', 'MATQM230RSBR', 'MATQM430RLWH',
    'MATQM730RZOM', 'MATQP511RKWH', 'MATQP512RKWH', 'MATQP513RKWH', 'MATQP514RKWH', 'MATSF520RKIV', 'MATSM430RLWH',
    'WPUTDC10RRNW', 'WPUTDC114RNW', 'WPUTDF104RNW', 'WPUTDF114RNW'}


class RealRun:
    """run과 같은 순서로 로드·계산·대조·검증한 결과."""

    def __init__(self):
        wb = openpyxl.load_workbook(BASE, data_only=True)
        actual_wb = openpyxl.load_workbook(ACTUAL, data_only=True)
        base = load_base(wb)  # 통합문서는 한 번만 읽는다
        self.results = base.results(CalcSettings())
        self.rows = compare_with_excel(self.results, base.summaries)
        self.verdicts = sku_verdicts(self.rows)
        # 산식 재현 여부는 계산값끼리 본다: 입력값인 현재고 대조를 뺀 판정
        self.formula_verdicts = sku_verdicts([r for r in self.rows if r.field != 'onhand'])
        self.actual_sales = load_sales(actual_wb)
        self.pm, self.sales, self.fcst = base.pm, base.sales, base.fcst
        self.onhand, self.corrections = base.onhand, base.corrections
        self.validation = validate(self.results, self.actual_sales, range(35, 44))
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
    # 현재고를 뺀 필드(수요·σ·LT·안전재고·ROP·일수)로 판정한다. 현재고 차이는 아래 테스트에서 따로 본다
    v = real_run.formula_verdicts
    weighted = {s: st for (b, s), st in v.items() if b == '가중평균'}
    assert len(weighted) == 101
    mismatched = {s for s, st in weighted.items() if st == '불일치'}
    # 산식으로 계산된 42개 SKU는 모두 일치, 불일치는 위 두 집합 안에서만
    formula_skus = {s for s in weighted if s not in COPIED_FROM_ACTUAL | FCST_BLANK}
    assert len(formula_skus) == 42
    assert {weighted[s] for s in formula_skus} == {'일치'}
    assert mismatched <= COPIED_FROM_ACTUAL | FCST_BLANK
    assert FCST_BLANK <= mismatched


def test_onhand_differences_surface(real_run):
    onhand = {r.sku for r in real_run.rows
              if r.basis == '가중평균' and r.field == 'onhand' and r.status in ('불일치', '계산값없음')}
    assert onhand == ONHAND_DIFF
    # 현재고까지 넣은 최종 판정의 불일치 = 산식 불일치 ∪ 현재고 차이
    final = {s for (b, s), st in real_run.verdicts.items() if b == '가중평균' and st == '불일치'}
    formula = {s for (b, s), st in real_run.formula_verdicts.items() if b == '가중평균' and st == '불일치'}
    assert final == formula | ONHAND_DIFF


def test_actual_and_fcst_basis_demand_match(real_run):
    # 엑셀에 값이 있는 행 기준: 실적기준 수요 59행 모두 일치.
    # FCST기준 46행 중 42행 일치, 불일치 4행은 전사본 FCST 칸이 비어 있는 SEN SKU
    assert real_run.demand_match('실적기준') == (59, 59)
    assert real_run.demand_match('FCST기준') == (42, 46)
    assert real_run.demand_mismatch_skus('FCST기준') == FCST_BLANK


def test_validation_summary_filled(real_run):
    overall = {m.basis: m for m in real_run.metrics if m.scope == '전체'}
    assert set(overall) == set(VALIDATION_BASES)
    # Task 7 결정: 실적·FCST 입력이 없는 SKU는 채점 제외(FCST없음 24, 실적없음+FCST없음 15, 실적파일없음 10) → 126 중 77.
    # BID SKU 코드를 BIDSS1DR15WH로 통일해 이제 실제 파일과 매칭된다(실적파일없음 11→10, 검증 SKU +1).
    # 따라서 실제 파일에만 있는 '계산대상없음' 행은 없다(단위 테스트에서 별도 검증)
    assert overall['가중평균'].n_demand >= 70 and overall['가중평균'].n_lt > 0 and overall['기존수기'].n_lt > 0
    assert [r.sku for r in real_run.validation if r.memo == '계산대상없음'] == []
    assert sum(r.memo == '실적파일없음' for r in real_run.validation) == 10
    # 보정계수이상 3개 SKU는 실적기준·FCST기준 채점에서 빠지고, 가중평균·기존수기 채점은 그대로다
    bad = {r.sku: r for r in real_run.validation if r.sku in {'FLTACLR11HWH', 'SENSAWR1KRDW', 'SENSAWR1KRNG'}}
    assert len(bad) == 3
    for r in bad.values():  # 다른 제외 사유(FCST없음 등)가 있으면 그 사유가 우선이라 어쨌든 채점되지 않는다
        for b in ('실적기준', 'FCST기준'):
            assert r.checks[b].excluded is not None and r.checks[b].covered is None, (r.sku, b)


def test_abnormal_correction_flagged(real_run):
    # 보정계수 0 또는 3 초과(예: SENSAWR1KRNG k=29.82)는 계산은 그대로 하고 메모만 붙인다
    memo = {r.inputs.sku: r.memo for r in real_run.results}
    assert {s for s, m in memo.items() if '보정계수이상' in m} == {'FLTACLR11HWH', 'SENSAWR1KRDW', 'SENSAWR1KRNG'}


@pytest.fixture(scope='module')
def real_grades():
    if not BASE.exists():
        pytest.skip('실데이터 파일이 없다')
    base = load_base(openpyxl.load_workbook(BASE, data_only=True))
    return list(base.pm.values()), base.grades


def test_grades_match_excel_all_rows(real_grades):
    rows, grades = real_grades
    assert len(rows) == 116
    # 재고등급은 116행 모두 엑셀과 같다
    assert all(r.grade == grades[r.sku].grade for r in rows)
    # 종합점수는 품절이력 C인 1행만 다르다: 기준 문서는 C=1점, 기존 엑셀은 C=0점을 썼다
    off = [r for r in rows if grade_verdict(r, grades[r.sku]) != '일치']
    assert len(off) == 1
    assert off[0].quality_grade == 'C' and grades[off[0].sku].total == off[0].excel_total + 1
    assert sum(g.turnover is None for g in grades.values()) == 16


def test_metrics_unchanged_with_computed_grades(real_run):
    # 계산 등급이 엑셀 등급과 모두 같으므로, 어느 쪽 등급으로 만들어도 검증 지표가 같다
    r = real_run

    def metrics(grades):
        inputs = build_inputs(r.pm, r.sales, r.fcst, r.onhand, r.corrections, CalcSettings(), grades=grades)
        return summarize(validate([compute_sku(i) for i in inputs], r.actual_sales, range(35, 44)))

    excel = metrics({s: row.grade for s, row in r.pm.items()})
    computed = metrics({s: g.grade for s, g in compute_grades(list(r.pm.values())).items()})
    assert excel == computed


def test_virtual_supplement_real(tmp_path):
    if not SUPPLEMENT.exists():
        pytest.skip('보완본 파일이 없다')
    out = tmp_path / 'v.xlsx'
    changes = build_virtual(SUPPLEMENT, out)
    assert any(c.reason.startswith('회전율 분모 0') for c in changes)
    wb = openpyxl.load_workbook(out, data_only=True)
    base = load_base(wb)
    assert len(base.pm) == 116
    assert {grade_verdict(r, base.grades[r.sku]) for r in base.pm.values()} == {'일치'}
    rows = compare_with_excel(base.results(CalcSettings()), base.summaries)
    # 수요 0인 SKU는 일수 계산식이 0으로 나눠져 계산·엑셀 모두 빈 칸(엑셀값없음, 채울 값 자체가 없다)
    assert {r.status for r in rows} <= {'일치', '계산값없음', '엑셀값없음'}
    assert all(r.calc is None and r.field in ('target_days', 'ss_days') for r in rows if r.status == '엑셀값없음')
    assert {r.sku for r in rows if r.status == '계산값없음'} == {'FLTACPP240WH', 'FLTACL140PWH'}


def test_weight_search_real():
    from safety_stock.weight_search.search import run_search
    base = VIRTUAL if VIRTUAL.exists() else SUPPLEMENT
    if not base.exists():
        pytest.skip('가상 보완본과 보완본이 모두 없다')
    r = run_search(base, ACTUAL, CalcSettings(), range(35, 39), range(39, 44), range(35, 44))
    assert {c.group for c in r.choices} >= {'FLT', 'MAT'}
    assert {c.method for c in r.choices if c.group in ('ACL', 'BID')} == {'공통최적'}
    assert set(r.metrics) == {('검증', '현재'), ('검증', '탐색'), ('전체구간', '현재'), ('전체구간', '탐색')}
