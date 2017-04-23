import time
import cProfile
import os

import xlsxwriter
from openpyxl import Workbook

def writer():
    start = time.clock()
    workbook = xlsxwriter.Workbook('sample2.xlsx',
                                   options={'constant_memory':True})
    worksheet = workbook.add_worksheet()
    for i in range(0, 10000):
        worksheet.write_row(i, 0, range(0, 50))
    workbook.close()
    end = time.clock()
    print("xlsxwriter {0}s".format(end - start))


def pixel():
    start = time.clock()
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()

    for i in range(0, 10000):
        ws.append(range(0, 50))

    wb.save("sample.xlsx")
    end = time.clock()
    print("openpyxl {0}s".format(end - start))

def cleanup():
    files = ("sample.xlsx", "sample2.xlsx")
    for fn in files:
        if os.path.exists(fn):
            os.remove(fn)

if __name__ == '__main__':
    writer()
    pixel()
    #cProfile.run("writer()", sort="tottime")
    #cProfile.run("pixel()", sort="tottime")
    cleanup()
