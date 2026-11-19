"""결과 xlsx 작성과 전체 실행(run). 쓰기는 이 모듈에서만 한다."""
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from safety_stock.common.xlsx import InputError
from safety_stock.inventory_days.calc import (
    BASES,
    CORRECTION_MAX,
    CalcSettings,
    SkuResult,
    build_inputs,
    compute_sku,
)
from safety_stock.inventory_days.compare import CompareRow, compare_with_excel, sku_verdicts
from safety_stock.inventory_days.loader import load_onhand, load_sales, load_summary
from safety_stock.inventory_days.validate import (
    VALIDATION_BASES,
    SummaryMetric,
    ValidationRow,
    summarize,
    validate,
)
from safety_stock.mp_simulation.loader import load_weekly_plan
from safety_stock.purchase_master.loader import load_purchase_master
from safety_stock.purchase_master.model import Z_BY_GRADE

REPORT_SHEETS = ['산출결과', '엑셀대조', '검증_SKU', '검증요약', '적용기준']

QTY, DAYS, PCT = '#,##0', '0.0', '0.0%'
SIGMA = '#,##0.0'  # σ류는 작은 값이 0으로 보이지 않도록 소수 1자리

# 표시 단계에서만 쓰는 한국어 이름·라벨
FIELD_LABELS = {
    'demand': '수요', 'sigma': 'σ', 'lt_demand': 'LT평균수요', 'lt_sigma': 'LT수요편차',
    'ss': '안전재고', 'rop': 'ROP', 'target_days': '적정재고일수', 'ss_days': '안전재고일수',
    'onhand': '현재고',
}
ORDER_LABELS = {'ORDER': '발주필요', 'OK': '정상'}
COVER_LABELS = {True: '방어', False: '결품'}

# 산출결과의 기준별 열: (제목, BasisResult 속성, 표시 형식)
_BASIS_COLS = [
    ('수요', 'demand_week', QTY), ('σ', 'sigma_week', SIGMA),
    ('LT평균수요', 'lt_demand', QTY), ('LT수요편차', 'lt_sigma', SIGMA),
    ('Z', 'z', '0.000'), ('안전재고', 'safety_stock', QTY), ('ROP', 'rop', QTY),
    ('적정재고일수', 'target_days', DAYS), ('안전재고일수', 'ss_days', DAYS),
    ('현재고일수', 'onhand_days', DAYS), ('발주', 'order', None),
]
_COMMON_COLS = [
    ('SKU', lambda i, r: i.sku, None), ('품목그룹', lambda i, r: i.group, None),
    ('품목명', lambda i, r: i.name, None), ('재고등급', lambda i, r: i.grade, None),
    ('실적가중치', lambda i, r: i.weight_actual, '0.00'), ('FCST가중치', lambda i, r: i.weight_fcst, '0.00'),
    ('리드타임(일)', lambda i, r: i.lead_time_days, DAYS), ('보정계수', lambda i, r: i.correction, '0.000'),
    ('현재고', lambda i, r: i.onhand, QTY),
    ('실적평균', lambda i, r: r.actual_mean, QTY), ('FCST평균', lambda i, r: r.fcst_mean, QTY),
]
_CHECK_COLS = [
    ('목표재고량', 'target_qty', QTY), ('예측주수요', 'predicted_week', QTY),
    ('수요오차', 'error_week', QTY), ('오차기여율', 'error_share', PCT), ('LT수요방어', 'covered', None),
    ('실제기준재고일수', 'actual_days', DAYS), ('제외사유', 'excluded', None),
]
_SUMMARY_COLS = [
    ('범위', 'scope', None), ('기준', 'basis', None), ('수요비교SKU수', 'n_demand', '0'),
    ('WAPE', 'wape', PCT), ('편향', 'bias', PCT), ('LT검증SKU수', 'n_lt', '0'),
    ('결품방어율', 'coverage', PCT), ('실제기준재고일수중앙값', 'median_actual_days', DAYS),
]


def _write_table(ws, headers, rows, formats):
    """제목 1행 + 데이터 행. 첫 행 고정, 자동필터, 제목 굵게."""
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=h).font = Font(bold=True)
        ws.column_dimensions[get_column_letter(c)].width = max(10, len(str(h)) * 2 + 2)
    for r, row in enumerate(rows, start=2):
        for c, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            fmt = formats[c - 1]
            if fmt and isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.number_format = fmt
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}{max(len(rows) + 1, 2)}'


def write_report(path, results: list[SkuResult], compare_rows: list[CompareRow],
                 validation_rows: list[ValidationRow], metrics: list[SummaryMetric],
                 meta: dict[str, str]) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheets = {name: wb.create_sheet(name) for name in REPORT_SHEETS}

    # 산출결과
    headers = [c[0] for c in _COMMON_COLS]
    formats = [c[2] for c in _COMMON_COLS]
    for b in BASES:
        headers += [f'{b}_{t}' for t, _, _ in _BASIS_COLS]
        formats += [f for _, _, f in _BASIS_COLS]
    headers += ['기존안전재고일수', '점검메모']
    formats += [DAYS, None]
    rows = []
    for res in results:
        inp = res.inputs
        row = [get(inp, res) for _, get, _ in _COMMON_COLS]
        for b in BASES:
            for _, attr, _ in _BASIS_COLS:
                v = getattr(res.bases[b], attr)
                row.append(ORDER_LABELS.get(v, v) if attr == 'order' else v)
        row += [inp.current_ss_days, ', '.join(res.memo)]
        rows.append(row)
    _write_table(sheets['산출결과'], headers, rows, formats)

    # 엑셀대조: 필드별 행 + 그 SKU의 판정(일치/불일치/비교불가)
    verdicts = sku_verdicts(compare_rows)
    rows = [[r.sku, r.basis, FIELD_LABELS.get(r.field, r.field), r.calc, r.excel, r.diff, r.status, verdicts[(r.basis, r.sku)]]
            for r in compare_rows]
    _write_table(sheets['엑셀대조'],
                 ['SKU', '기준', '필드', '계산값', '엑셀값', '차이', '상태', 'SKU판정'], rows,
                 [None, None, None, '#,##0.00', '#,##0.00', '#,##0.00', None, None])

    # 검증_SKU: 기준별로 채점 제외 사유도 함께
    headers = ['SKU', '품목그룹', '리드타임(일)', '실제주평균', '실제LT수요', '메모']
    formats = [None, None, DAYS, QTY, QTY, None]
    for b in VALIDATION_BASES:
        headers += [f'{b}_{t}' for t, _, _ in _CHECK_COLS]
        formats += [f for _, _, f in _CHECK_COLS]
    rows = []
    for v in validation_rows:
        row = [v.sku, v.group, v.lead_time_days, v.actual_week_avg, v.actual_lt_demand, v.memo]
        for b in VALIDATION_BASES:
            for _, attr, _ in _CHECK_COLS:
                val = getattr(v.checks[b], attr)
                row.append(COVER_LABELS.get(val, val) if attr == 'covered' else val)
        rows.append(row)
    _write_table(sheets['검증_SKU'], headers, rows, formats)

    # 검증요약: 값이 None인 지표는 빈 칸
    rows = [[getattr(m, attr) for _, attr, _ in _SUMMARY_COLS] for m in metrics]
    _write_table(sheets['검증요약'], [h for h, _, _ in _SUMMARY_COLS], rows,
                 [f for _, _, f in _SUMMARY_COLS])

    # 적용기준
    _write_table(sheets['적용기준'], ['항목', '내용'], [[k, v] for k, v in meta.items()], [None, None])
    sheets['적용기준'].column_dimensions['A'].width = 24
    sheets['적용기준'].column_dimensions['B'].width = 70

    Path(path).parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(path)
    except PermissionError:
        raise InputError(f'결과 파일을 저장할 수 없습니다(엑셀에서 열려 있으면 닫고 다시 실행): {path}') from None


def _open(path: Path):
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except FileNotFoundError:
        raise InputError(f'파일을 찾을 수 없습니다: {path}') from None
    except Exception as e:  # 손상됐거나 xlsx가 아닌 파일
        raise InputError(f'엑셀 파일을 열 수 없습니다: {path} ({e})') from e


def run(base: Path, actual: Path, out: Path, settings: CalcSettings, validation_weeks: range) -> dict:
    """기준 엑셀·실제 출고 엑셀을 읽어 계산·대조·검증하고 결과 xlsx를 쓴다."""
    base, actual, out = Path(base), Path(actual), Path(out)
    for src in (base, actual):
        if out.resolve() == src.resolve():
            raise InputError(f'결과 파일 경로가 입력 파일과 같습니다: {out}')
    wb = _open(base)
    actual_wb = _open(actual)

    pm = load_purchase_master(wb)
    sales = load_sales(wb)
    fcst = load_weekly_plan(wb, 'SCM FCST')
    onhand = load_onhand(wb)
    weighted_summary = load_summary(wb, 'Summary_가중평균', required=('correction',))
    corrections = {sku: row.get('correction') for sku, row in weighted_summary.items()}
    summaries = {
        '가중평균': weighted_summary,
        '실적기준': load_summary(wb, 'Summary_실적'),
        'FCST기준': load_summary(wb, 'Summary_FCST'),
    }

    results = [compute_sku(i) for i in build_inputs(pm, sales, fcst, onhand, corrections, settings)]
    compare_rows = compare_with_excel(results, summaries)
    validation_rows = validate(results, load_sales(actual_wb), validation_weeks)
    metrics = summarize(validation_rows)

    used_labels = fcst.week_labels[:max(settings.fcst_weeks, 0)]
    fcst_range = f'{used_labels[0]}~{used_labels[-1]}' if used_labels else '주차 없음'
    meta = {
        '실행 시각': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        '기준 파일': base.name,
        '실제 출고 파일': actual.name,
        '실적 주차': f'WK{settings.actual_first_week:02d}~WK{settings.actual_last_week:02d}',
        'FCST 주 수': f'앞 {settings.fcst_weeks}주 {fcst_range} ({fcst.sheet})',
        '검증 주차': f'WK{validation_weeks.start:02d}~WK{validation_weeks.stop - 1:02d}',
        'Z값': ', '.join(f'{g}={z}' for g, z in Z_BY_GRADE.items()),
        '평균·σ 정의': '평균은 AVERAGE, σ는 STDEV.P(모표준편차). 실적 빈 칸은 제외하고 계산, FCST 빈 칸은 0으로 보고 주 수 전체로 계산',
        '가중평균 수요': '실적평균×실적가중치 + FCST평균×FCST가중치',
        '가중평균 σ': '실적σ×실적가중치 + FCSTσ×FCST가중치',
        '실적기준 수요': '실적평균×보정계수',
        'FCST기준 수요': 'FCST평균×보정계수',
        '안전재고': 'Z × σ × √(리드타임/7)',
        'ROP': 'LT평균수요 + 안전재고',
        '일수 계산': '적정재고일수·안전재고일수·현재고일수 = 값 ÷ (주수요/7)',
        '발주 규칙': '현재고 ≤ ROP이면 발주필요, 아니면 정상',
        '대조 항목': '수요, σ, LT평균수요, LT수요편차, 안전재고, ROP, 적정재고일수, 안전재고일수, 현재고'
                 '(현재고는 기준과 무관한 값이라 세 기준 행 모두에 같은 비교가 들어간다)',
        '대조 허용오차': '엑셀 표시 자릿수 반올림 ±0.5단위 + 엑셀값의 0.5%(보정계수·가중치 반올림 여유)',
        '보정계수 이상': f'보정계수가 0이거나 {CORRECTION_MAX:g}보다 크면 점검메모 보정계수이상(계산은 그대로)',
        '검증 원칙': '실제 출고는 검증에만 쓰고 계산 입력에는 쓰지 않는다',
        '검증 대상 SKU': '기준 파일과 실제 출고 파일 두 파일에 모두 있는 SKU. '
                     '실제 파일에 없으면 메모 실적파일없음, 실제 파일에만 있으면 메모 계산대상없음(채점 안 함)',
        '기존수기 목표재고량': '기존 안전재고 값(일) × 가중평균 일평균수요(가중평균 주수요/7)',
        '산식 기준 목표재고량': 'ROP',
        '결품 방어': f'목표재고량 ≥ 실제 LT 구간 수요이면 방어. 리드타임 {7 * len(validation_weeks)}일'
                 '(검증 구간) 이하 SKU만',
        '오차기여율': '|수요오차| ÷ 그 기준 전체 Σ|수요오차| (범위: 전체)',
        '범위 정의': "'전체'=기준별로 채점 가능한 SKU 모두(기준마다 SKU 수가 다를 수 있음). "
                 "'공통'=수요 지표는 세 산식 기준 모두 수요오차가 있는 SKU, LT 지표는 네 기준 모두 방어 여부가 있는 SKU. "
                 "품목그룹=공통 범위의 SKU를 그 그룹으로 거른 것",
    }
    write_report(out, results, compare_rows, validation_rows, metrics, meta)

    verdicts = sku_verdicts(compare_rows)
    return {
        'sku_count': len(results),
        'match_count_weighted': sum(1 for (b, _), st in verdicts.items() if b == '가중평균' and st == '일치'),
        'out': out,
    }
