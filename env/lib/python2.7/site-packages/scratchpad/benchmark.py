##############################################################################
#
# Simple Python program to benchmark several Python Excel writing modules.
#
# python bench_excel_writers.py [num_rows] [num_cols]
#
# Copyright 2013-2015, John McNamara, jmcnamara@cpan.org
#

import sys
from time import clock

import openpyxl
import xlsxwriter


# Default to 1000 rows x 50 cols.
if len(sys.argv) > 1:
    row_max = int(sys.argv[1])
    col_max = 50
else:
    row_max = 1000
    col_max = 50

if len(sys.argv) > 2:
    col_max = int(sys.argv[2])


def print_elapsed_time(module_name, elapsed):
    """ Print module run times in a consistent format. """
    print("    %-22s: %6.2f" % (module_name, elapsed))


def time_xlsxwriter():
    """ Run XlsxWriter in default mode. """
    start_time = clock()

    workbook = xlsxwriter.Workbook('xlsxwriter.xlsx')
    worksheet = workbook.add_worksheet()

    for row in range(0, row_max, 2):
        string_data = ["Row: %d Col: %d" % (row, col) for col in range(col_max)]
        worksheet.write_row(row, 0, string_data)

        num_data = [row + col for col in range(col_max)]
        worksheet.write_row(row + 1, 0, num_data)

    workbook.close()

    elapsed = clock() - start_time
    print_elapsed_time('xlsxwriter', elapsed)


def time_xlsxwriter_optimised():
    """ Run XlsxWriter in optimised/constant memory mode. """
    start_time = clock()

    workbook = xlsxwriter.Workbook('xlsxwriter_opt.xlsx',
                                   options={'constant_memory': True})
    worksheet = workbook.add_worksheet()

    for row in range(0, row_max, 2):
        string_data = ["Row: %d Col: %d" % (row, col) for col in range(col_max)]
        worksheet.write_row(row, 0, string_data)

        num_data = [row + col for col in range(col_max)]
        worksheet.write_row(row + 1, 0, num_data)

    workbook.close()

    elapsed = clock() - start_time
    print_elapsed_time('xlsxwriter (optimised)', elapsed)


def time_openpyxl():
    """ Run OpenPyXL in default mode. """
    start_time = clock()

    workbook = openpyxl.workbook.Workbook()
    worksheet = workbook.active

    for row in range(row_max // 2):

        string_data = ("Row: %d Col: %d" % (row, col) for col in range(col_max))
        worksheet.append(string_data)

        num_data = (row + col for col in range(col_max))
        worksheet.append(num_data)

    workbook.save('openpyxl.xlsx')

    elapsed = clock() - start_time
    print_elapsed_time('openpyxl', elapsed)


def time_openpyxl_optimised():
    """ Run OpenPyXL in optimised mode. """
    start_time = clock()

    workbook = openpyxl.workbook.Workbook(write_only=True)
    worksheet = workbook.create_sheet()

    for row in range(row_max // 2):
        string_data = ("Row: %d Col: %d" % (row, col) for col in range(col_max))
        worksheet.append(string_data)

        num_data = (row + col for col in range(col_max))
        worksheet.append(num_data)

    workbook.save('openpyxl_opt.xlsx')

    elapsed = clock() - start_time
    print_elapsed_time('openpyxl   (optimised)', elapsed)


print("")
print("Versions:")
print("    %-12s: %s" % ('python', sys.version[:5]))
print("    %-12s: %s" % ('openpyxl', openpyxl.__version__))
print("    %-12s: %s" % ('xlsxwriter', xlsxwriter.__version__))
print("")

print("Dimensions:")
print("    Rows = %d" % row_max)
print("    Cols = %d" % col_max)
print("")

print("Times:")
time_xlsxwriter_optimised()
time_xlsxwriter()
time_openpyxl_optimised()
time_openpyxl()
print("")
