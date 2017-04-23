from io import BytesIO
from lxml.etree import xmlfile
import os
from random import randint

from openpyxl import Workbook
# from openpyxl.xml.functions import XMLGenerator

def make_worksheet():
    wb = Workbook()
    ws = wb.active
    for i in range(1000):
        ws.append(list(range(100)))
    return ws


def lxml_writer(ws=None):
    from openpyxl.writer.etree_worksheet import write_rows
    if ws is None:
        ws = make_worksheet()

    out = BytesIO()
    with xmlfile(out) as xf:
        write_rows(xf, ws)
    #with open("lxml_writer.xml", "wb") as dump:
        #dump.write(out.getvalue())
    #ws.parent.save("lxml_writer.xlsx")


def make_dump_worksheet():
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()
    return ws

def dump_writer(ws=None):
    if ws is None:
        ws = make_dump_worksheet()
    for i in range(1000):
        ws.append(list(range(100)))


COLUMNS = 100
ROWS = 1000
BOLD = 1
ITALIC = 2
UNDERLINE = 4
RED_BG = 8
formatData = [[None] * COLUMNS for _ in range(ROWS)]

def generate_format_data():
    for row in range(ROWS):
        for col in range(COLUMNS):
            formatData[row][col] = randint(1, 15)


def styled_sheet():
    from openpyxl import Workbook
    from openpyxl.styles import Font, Style, Fill, Color, colors

    wb = Workbook()
    ws = wb.active
    ws.title = 'Test 1'

    red_fill = Fill()
    red_fill.fill_type = 'solid'
    red_fill.start_color = Color(Color.RED),
    red_fill.end_color=Color(Color.RED)
    empty_fill = Fill()
    styles = []
    # pregenerate relevant styles
    for row in range(ROWS):
        _row = []
        for col in range(COLUMNS):
            cell = ws.cell(row=row+1, column=col+1)
            cell.value = 1
            font = Font()
            fill = Fill()
            if formatData[row][col] & BOLD:
                font.bold = True
            if formatData[row][col] & ITALIC:
                font.italic = True
            if formatData[row][col] & UNDERLINE:
                font.underline = 'single'
            if formatData[row][col] & RED_BG:
                fill = red_fill
            style = Style()
            style.font = font
            style.fill = fill
            ws._styles[cell.address] = style


    #wb.save(get_output_path('test_openpyxl_style_std_pregen.xlsx'))


def read_workbook():
    from openpyxl import load_workbook
    folder = os.path.split(__file__)[0]
    src = os.path.join(folder, "files", "very_large.xlsx")
    wb = load_workbook(src)
    return wb


def rows(wb):
    ws = wb.active
    rows = ws.iter_rows()
    for r, row in enumerate(rows):
        for c, col in enumerate(row):
            pass
    print((r+1)* (c+1), "cells")


def col_index1():
    from openpyxl.cell import get_column_letter
    for i in range(1, 18279):
        c = get_column_letter(i)





"""
Sample use
import cProfile
ws = make_worksheet()
cProfile.run("profiling.lxml_writer(ws)",
             sort="tottime",
             filename="stats")
from pstats import Stats
st = Stats("stats")
st.print_callers("_bind_value")
st.print_callees("_bind_value")
"""


if __name__ == '__main__':
    import cProfile
    #ws = make_worksheet()
    #wb = read_workbook()
    #cProfile.run("rows(wb)", sort="tottime")
    #cProfile.run("make_worksheet()", sort="tottime")
    #cProfile.run("lxml_writer(ws)", sort="tottime")
    # generate_format_data()
    # cProfile.run("styled_sheet()", sort="tottime")
    ws = make_dump_worksheet()
    cProfile.run("dump_writer(ws)", sort="tottime")
    # cProfile.run("col_index1()", sort="tottime")
