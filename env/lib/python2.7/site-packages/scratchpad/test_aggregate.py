from openpyxl import Workbook

import pytest
from openpyxl.styles import Font


@pytest.fixture
def dummy_workbook():
    """
    Creates a dummy worksheet 5 x 5 with values 0-24
    """
    sz = 5
    wb = Workbook()
    ws = wb.active
    for i in range(0, sz):
        for j in range(0, sz):
            ws.cell(row=i+1, column=j+1).value = i*sz + j
    return wb


@pytest.fixture
def Proxy():

    from ..aggregate import Proxy
    return Proxy


class TestProxy:


    def test_move_row_down(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        assert ws.max_row == 5
        proxy = Proxy(ws)

        proxy._move_cells(min_row=5, offset=1, attr="row")

        assert ws.max_row == 6
        assert [c.value for c in ws[5]] == [None]*5

    def test_move_col_right(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        assert ws.max_column == 5
        proxy = Proxy(ws)

        proxy._move_cells(min_col=3, offset=2, attr="col_idx")

        assert ws.max_column == 7
        assert [c.value for c in ws['D']] == [None]*5

    def test_move_row_up(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        assert ws.max_column == 5
        proxy = Proxy(ws)

        proxy._move_cells(min_row=4, offset=-1, attr="row")

        assert ws.max_row
        assert [c.value for c in ws['A']] == [0, 5, 15, 20]


    def test_insert_rows(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)

        proxy.insert_rows(2)
        assert [c.value for c in ws[2]] == [None]*5


    def test_insert_cols(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)

        proxy.insert_cols(3)
        assert ws.max_column == 6
        assert [c.value for c in ws['D']] == [2, 7, 12, 17, 22]


    def test_delete_rows(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)

        proxy.delete_rows(2)
        assert ws.max_row == 4
        assert [c.value for c in ws[1]] == [0, 1, 2, 3, 4]
        assert [c.value for c in ws[2]] == [10, 11, 12, 13, 14]


    def test_delete_cols(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)

        proxy.delete_cols(3)
        assert ws.max_column == 4
        assert [c.value for c in ws[3]] == [10, 11, 13, 14]


    def test_fill_rows(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)

        rows = [
        [10]*5,
        ]

        proxy.fill_rows(2, rows)
        assert [c.value for c in ws[2]] == [10]*5


    def test_fill_cols(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)

        cols = [
        [5]*5,
        ]

        proxy.fill_cols(1, cols)
        assert [c.value for c in ws['A']] == [5]*5


    def test_insert_row_values(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)

        rows = [
        [4]*3,
        [6]*3
        ]

        proxy.insert(5, rows)
        assert ws.max_row == 7
        assert [c.value for c in ws[6]] == [6]*3 + [None]*2


    def test_insert_col_values(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)

        cols = [
        [4]*3,
        [6]*3
        ]

        proxy.insert(5, cols, dimension="column")
        assert ws.max_column == 7
        assert [c.value for c in ws['F']] == [6]*3 + [None]*2


    def test_swap_consistency(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)
        c1_row = 2
        c2_row = 5
        c1_cell = ws[c1_row][0]
        c2_cell = ws[c2_row][0]
        proxy.swap_cells(ws[c1_row], ws[c2_row])
        assert c2_cell.row == c1_row
        assert c1_cell.row == c2_row


    def test_swap_rows(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)

        proxy.swap_cells(ws[1], ws[3])
        proxy.swap_cells(ws[2], ws[6])
        assert [c.value for c in ws[1]] == [10, 11, 12, 13, 14]
        assert [c.value for c in ws[3]] == [0, 1, 2, 3, 4]
        assert [c.value for c in ws[6]] == [5, 6, 7, 8, 9]


    def test_swap_cols(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)

        proxy.swap_cells(ws['A'], ws['D'])
        assert [c.value for c in ws['A']] == [3, 8, 13, 18, 23]
        assert [c.value for c in ws['D']] == [0, 5, 10, 15, 20]


    def test_clear_cells(self, Proxy, dummy_workbook):
        ws = dummy_workbook.active
        proxy = Proxy(ws)

        proxy.clear(min_row=2, min_col=2, max_row=4, max_col=4)

        assert [c.value for c in ws['B']] == [1, None, None, None, 21]
