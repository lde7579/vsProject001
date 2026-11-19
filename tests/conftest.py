import pytest
from openpyxl import Workbook


def _make_wb(sheets):
    """시트 이름 -> 행 목록(1행부터)으로 메모리 통합문서를 만든다."""
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r, row in enumerate(rows, start=1):
            for c, v in enumerate(row, start=1):
                if v is not None:
                    ws.cell(row=r, column=c, value=v)
    return wb


@pytest.fixture
def make_wb():
    return _make_wb
