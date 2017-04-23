import re
from string import ascii_letters
from openpyxl import Workbook

import pytest


@pytest.fixture
def dummy_worksheet():
    """
    Creates a dummy worksheet 10 x 10 with values 0-24
    """
    sz = 5
    text = list(ascii_letters)*2
    wb = Workbook()
    ws = wb.active
    for i in range(0, 10):
        for j in range(0, 10):
            ws.cell(row=i+1, column=j+1, value=text.pop())
    wb.save("letters.xlsx")
    return ws


@pytest.fixture
def Indexer():

    from ..indexer import Indexer
    return Indexer


class TestIndexer:


    def test_ctor(self, Indexer, dummy_worksheet):
        idx = Indexer(dummy_worksheet)

        assert idx._db == {
            'A': [(3, 6), (8, 8)],
            'B': [(3, 5), (8, 7)],
            'C': [(3, 4), (8, 6)],
            'D': [(3, 3), (8, 5)],
            'E': [(3, 2), (8, 4)],
            'F': [(3, 1), (8, 3)],
            'G': [(2, 10), (8, 2)],
            'H': [(2, 9), (8, 1)],
            'I': [(2, 8), (7, 10)],
            'J': [(2, 7), (7, 9)],
            'K': [(2, 6), (7, 8)],
            'L': [(2, 5), (7, 7)],
            'M': [(2, 4), (7, 6)],
            'N': [(2, 3), (7, 5)],
            'O': [(2, 2), (7, 4)],
            'P': [(2, 1), (7, 3)],
            'Q': [(1, 10), (7, 2)],
            'R': [(1, 9), (7, 1)],
            'S': [(1, 8), (6, 10)],
            'T': [(1, 7), (6, 9)],
            'U': [(1, 6), (6, 8)],
            'V': [(1, 5), (6, 7)],
            'W': [(1, 4), (6, 6)],
            'X': [(1, 3), (6, 5)],
            'Y': [(1, 2), (6, 4)],
            'Z': [(1, 1), (6, 3)],
            'a': [(6, 2)],
            'b': [(6, 1)],
            'c': [(5, 10)],
            'd': [(5, 9)],
            'e': [(5, 8), (10, 10)],
            'f': [(5, 7), (10, 9)],
            'g': [(5, 6), (10, 8)],
            'h': [(5, 5), (10, 7)],
            'i': [(5, 4), (10, 6)],
            'j': [(5, 3), (10, 5)],
            'k': [(5, 2), (10, 4)],
            'l': [(5, 1), (10, 3)],
            'm': [(4, 10), (10, 2)],
            'n': [(4, 9), (10, 1)],
            'o': [(4, 8), (9, 10)],
            'p': [(4, 7), (9, 9)],
            'q': [(4, 6), (9, 8)],
            'r': [(4, 5), (9, 7)],
            's': [(4, 4), (9, 6)],
            't': [(4, 3), (9, 5)],
            'u': [(4, 2), (9, 4)],
            'v': [(4, 1), (9, 3)],
            'w': [(3, 10), (9, 2)],
            'x': [(3, 9), (9, 1)],
            'y': [(3, 8), (8, 10)],
            'z': [(3, 7), (8, 9)]
        }


    def test_find(self, Indexer, dummy_worksheet):
        idx = Indexer(dummy_worksheet)
        search = re.compile('a', re.IGNORECASE)

        coords = idx.find(search)
        assert set(coords) ==  {'B6', 'F3', 'H8'}


    def test_find_all(self, Indexer, dummy_worksheet):
        idx = Indexer(dummy_worksheet)
        search = re.compile('a', re.IGNORECASE)

        assert idx.find_all(search) == [(3, 6), (6, 2), (8, 8)]


    def test_find_next(self, Indexer, dummy_worksheet):
        idx = Indexer(dummy_worksheet)
        search = re.compile('a', re.IGNORECASE)

        assert idx.find_next(search, min_row=3) == (6, 2)
