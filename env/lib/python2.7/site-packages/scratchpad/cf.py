from openpyxl import Workbook
from openpyxl import __version__
from openpyxl.formatting.rule import  CellIsRule
from openpyxl.styles import  Font, PatternFill
import timeit
from copy import copy

from cProfile import run

wb = Workbook()
ws = wb.worksheets[0]

pink_text = Font(color='9C0006')
green_text = Font(color='006100')
pink_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
purple_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

from openpyxl.utils import get_column_letter

r1 = CellIsRule(operator='greaterThan', font=pink_text, formula=[],
                fill=pink_fill
                )

r2 = CellIsRule(operator='lessThanOrEqual', formula=[],
                font=green_text,
                fill=purple_fill)

r3 = CellIsRule(operator='greaterThanOrEqual',                 font=green_text,
                fill=purple_fill, formula=[],
                )

r4 = CellIsRule(operator='lessThan',               font=pink_text,
              fill=pink_fill, formula=[],
              )

r5 = CellIsRule(operator='greaterThan', formula=['0'],
                font=pink_text,
                fill=pink_fill
              )

r6 = CellIsRule(operator='lessThan', formula=['0'],
               font=green_text,
               fill=purple_fill
               )

r7 = CellIsRule(operator='greaterThan', formula=['0'],
                font=green_text,
                fill=purple_fill
                )

r8 = CellIsRule(operator='lessThan', formula=['0'],
                font=pink_text,
                fill=pink_fill
                )

cf = ws.conditional_formatting

for row in range(1,500,10):
    for col in range(1,500,20):
        cell = '{0}{1}'.format(get_column_letter(col+4), row)
        rule = copy(r1)
        formula = ["{0}{1}".format(get_column_letter(col+6), row)]
        rule.formula = formula
        cf.add(cell, rule)

        rule = copy(r2)
        rule.formula = formula
        cf.add(cell, rule)


        cell = '{0}{1}'.format(get_column_letter(col+5), row)
        rule = copy(r3)
        formula = ["{0}{1}".format(get_column_letter(col+7), row)]
        rule.formula = formula
        cf.add(cell, rule)

        rule = copy(r4)
        rule.formula = formula
        cf.add(cell, rule)


        cell = '{0}{1}'.format(get_column_letter(col+17), row)
        rule = copy(r5)
        cf.add(cell, rule)

        rule = copy(r6)
        cf.add(cell, rule)


        cell = '{0}{1}'.format(get_column_letter(col+18), row)
        rule = copy(r7)
        cf.add(cell, rule)

        rule = copy(r8)
        cf.add(cell, rule)


def save_me():
  wb.save('test_{0}.xlsx'.format(__version__))

a = timeit.default_timer()
save_me()
print(timeit.default_timer()-a)
