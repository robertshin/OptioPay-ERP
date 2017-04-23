"""
Skeleton implementation of a proxy object for aggregate functions
"""

from operator import attrgetter
from itertools import zip_longest

from openpyxl.cell import Cell


class Proxy(object):


    def __init__(self, ws):
        self.ws = ws


    def insert_rows(self, idx, amount=1):
        """
        Insert rows before row==idx
        """
        self._move_cells(min_row=idx, offset=amount, attr="row")


    def insert_cols(self, idx, amount=1):
        """
        Insert columns before col==idx
        """
        self._move_cells(min_col=idx, offset=amount, attr="col_idx")


    def delete_rows(self, idx, amount=1):
        """
        Delete rows from row==idx
        """
        self._move_cells(min_row=idx+amount, offset=-amount, attr="row")


    def delete_cols(self, idx, amount=1):
        """
        Delete columns from col==idx
        """
        self._move_cells(min_col=idx+amount, offset=-amount, attr="col_idx")


    def insert(self, idx, seq, dimension="row"):
        """
        Insert a sequence of rows or columns at idx
        """
        amount = len(seq)
        if dimension == "row":
            self.insert_rows(idx, amount)
            self.fill_rows(idx, seq)
        elif dimension == "column":
            self.insert_cols(idx, amount)
            self.fill_cols(idx, seq)


    def fill_rows(self, idx, seq):
        """
        Overwrite rows starting at idx
        """
        for r_idx, row in enumerate(seq, idx):
            for c_idx, v in enumerate(row, 1):
                self.ws.cell(row=r_idx, column=c_idx, value=v)


    def fill_cols(self, idx, seq):
        """
        Overwrite columns
        """
        for c_idx, col in enumerate(seq, idx):
            for r_idx, v in enumerate(col, 1):
                self.ws.cell(row=r_idx, column=c_idx, value=v)


    def _move_cells(self, min_row=None, min_col=None, offset=0, attr=None):
        """
        Move cells by row or column
        """

        reverse = offset > 0 # start at the end if moving down

        all_cells = self.ws._cells

        cells = sorted(all_cells.values(), key=attrgetter(attr), reverse=reverse)

        for cell in cells:
            if min_row and cell.row < min_row:
                continue
            elif min_col and cell.col_idx < min_col:
                continue

            del all_cells[(cell.row, cell.col_idx)] # remove old ref

            val = getattr(cell, attr)
            setattr(cell, attr, val+offset) # calculate new coords

            all_cells[(cell.row, cell.col_idx)] = cell # add new ref


    def swap_cells(self, source, target):
        """
        Transpose two rows or two columns
        """
        all_cells = self.ws._cells
        for c1, c2 in zip_longest(source, target):
            key2 = (c2.row, c2.col_idx)
            key1 = (c1.row, c1.col_idx)
            all_cells[key2] = c1
            all_cells[key1] = c2
            c1.row, c1.col_idx = key2
            c2.row, c2.col_idx = key1


    def clear(self, min_row=1, min_col=1, max_row=None, max_col=None):
        """
        Remove a range of cells from the worksheet but do not move any cells
        """
        max_row = max_row or self.ws.max_row
        max_col = max_col or self.ws.max_column

        for row, col in sorted(self.ws._cells):
            if (min_row <= row <= max_row
                and min_col <= col <= max_col):
                del self.ws._cells[(row, col)]

