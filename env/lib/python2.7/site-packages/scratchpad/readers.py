"""
Compare read performance with xlrd
"""

import cProfile
from time import clock

test_file = "Issues/bug494.xlsx"
from openpyxl import load_workbook
from xlrd import open_workbook


def pixel():
    start = clock()
    counter = 0
    wb = load_workbook(test_file, read_only=True)
    for ws in wb:
        for row in ws.iter_rows():
            for cell in row:
                cell.value
    end = clock()
    print("openpyxl {0}s".format(end - start))


def xlrd():
    start = clock()
    counter = 0

    wb = open_workbook(test_file)
    for ws in wb.sheets():
        for idx in range(ws.nrows):
            row = ws.row(idx)
            for cell in row:
                cell.value

    end = clock()
    print("xlrd {0}s".format(end - start))


if __name__ == "__main__":
    xlrd()
    pixel()
    #cProfile.run("pixel()", sort="tottime")
    #cProfile.run("xlrd()", sort="tottime")
    #cProfile.run("checker()", sort="tottime")
