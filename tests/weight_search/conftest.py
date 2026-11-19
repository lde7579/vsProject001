import pytest


@pytest.fixture
def make_books(make_wb):
    """기준·실제 출고 엑셀을 임시 폴더에 저장하는 함수. 실제 출고는 WK35~WK43, 검증 구간(W39~43) 수량만 바꿀 수 있다."""
    def build(tmp_path, holdout_qty=20, tag=''):
        weeks = [f'WK{n:02d}' for n in range(1, 32)]
        fcst_labels = [f'W{n}(08/24)' for n in range(35, 48)]
        summary_head = ['SKU', '평균수요', '과소예측보정계수']
        base = make_wb({
            'W30 구매마스터': [
                ['SKU', '품목그룹', '품목명', '재고등급', '실적 가중치', 'FCST 가중치', '리드타임(일)', '안전재고 값',
                 '구매단가', '기초재고', '기말재고', 'FCST', '리드타임', '품질이슈', '리드타임', '품질이슈',
    '회전율', '단가', '회전율', '회전율', '단가'],
                ['AAA', 'FLT', '필터', 'A', 0.7, 0.3, 14, 30, 2000, 100, 100, 100, 'A', 'A', 3, 3],
                ['BBB', 'FLT', '필터2', 'B', 0.7, 0.3, 14, 30, 1000, 100, 100, 50, 'C', 'C', 3, 3],
            ],
            'Sales_WMS 출고': [['SKU', *weeks], ['AAA', *range(10, 41)], ['BBB', *([5] * 31)]],
            '0824_02 SCM FCST': [['품목 ID', 'x', *fcst_labels], ['AAA', None, *([20] * 13)],
                                 ['BBB', None, *([8] * 13)]],
            '0824_01 현재고': [['품목 ID', '현재고'], ['AAA', 100], ['BBB', 50]],
            'Summary_가중평균수요기준': [summary_head, ['AAA', 1, 1.2], ['BBB', 1, 1.2]],
            'Summary_실적기준': [summary_head, ['AAA', 1, 1.2]],
            'Summary_FCST기준': [summary_head, ['AAA', 1, 1.2]],
        })
        qty = [20] * 4 + [holdout_qty] * 5
        actual = make_wb({'출고실적': [['SKU', *[f'WK{n}' for n in range(35, 44)]],
                                   ['AAA', *qty], ['BBB', *([5] * 4), *([holdout_qty // 4] * 5)]]})
        base_path, actual_path = tmp_path / f'base{tag}.xlsx', tmp_path / f'actual{tag}.xlsx'
        base.save(base_path)
        actual.save(actual_path)
        return base_path, actual_path
    return build
