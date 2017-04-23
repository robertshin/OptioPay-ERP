import openpyxl
from random import randint
import time

import cProfile

ROWS = 1000
COLUMNS = 1000

testData = [[1] * COLUMNS] * ROWS
formatData = [[None] * COLUMNS for _ in range(ROWS)]

def generate_format_data():
    for row in range(ROWS):
        for col in range(COLUMNS):
            formatData[row][col] = randint(1, 15)


def run_openpyxl_optimised():
    stime = time.clock()
    wb = openpyxl.workbook.Workbook(optimized_write=True)
    ws = wb.create_sheet()
    ws.title = 'Test 1'
    for row in testData:
        ws.append(row)
    #wb.save("dump.xlsx")
    elapsed = time.clock() - stime
    print("openpyxl optimised value, %s, %s, %s" % (ROWS, COLUMNS, elapsed))


def run_openpyxl():
    stime = time.clock()
    wb = openpyxl.workbook.Workbook()
    ws = wb.create_sheet()
    ws.title = 'Test 1'
    for row in testData:
        ws.append(row)
    wb.save("dump.xlsx")
    elapsed = time.clock() - stime
    print("openpyxl standard value, %s, %s, %s" % (ROWS, COLUMNS, elapsed))


def single():
    from dispatch import Cell
    stime = time.clock()
    for row in testData:
        for value in row:
            c = Cell(None, "A", 1)
            c.value = value
    elapsed = time.clock() - stime


def standard():
    from openpyxl.cell import Cell
    stime = time.clock()
    for row in testData:
        for value in row:
            c = Cell(None, "A", 1)
            c.value = value
    elapsed = time.clock() - stime



if __name__ == "__main__":
    #cProfile.run("run_openpyxl_optimised()", sort="tottime")
    #cProfile.run("run_openpyxl()", sort="tottime")
    cProfile.run("standard()", sort="tottime")
    #cProfile.run("single()", sort="tottime")

    #run_openpyxl()
    #run_openpyxl_optimised()
