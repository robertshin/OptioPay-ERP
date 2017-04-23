import time
from numpy import random
from pandas import DataFrame

from openpyxl import Workbook

from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.styles.cell_style import StyleArray
from openpyxl.styles.named_styles import NamedStyle
from openpyxl.utils.dataframe import dataframe_to_rows

ft = Font(bold=True)
al = Alignment(horizontal="center")
side = Side(style="thin", color="000000")
border = Border(left=side, right=side, top=side, bottom=side)
highlight = NamedStyle(name="Pandas Title", font=ft, alignment=al, border=border)


def openpyxl_in_memory(df):
    """
    Import a dataframe into openpyxl
    """

    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(df, index=True, header=True):
        ws.append(r)

    for c in ws['A'] + ws['1']:
        c.style = 'Pandas'

    wb.save("pandas_openpyxl.xlsx")


from openpyxl.cell import WriteOnlyCell


def openpyxl_stream(df):
    """
    Write a dataframe straight to disk
    """
    wb = Workbook(write_only=True)
    ws = wb.create_sheet()

    cell = WriteOnlyCell(ws)
    cell.style = 'Pandas'

    def format_first_row(row, cell):

        for c in row:
            cell.value = c
            yield cell

    rows = dataframe_to_rows(df)
    first_row = format_first_row(next(rows), cell)
    ws.append(first_row)

    for row in rows:
        row = list(row)
        cell.value = row[0]
        row[0] = cell
        ws.append(row)

    wb.save("openpyxl_stream.xlsx")


def read_write(df1):
    """
    Create a worksheet from a Pandas dataframe and read it back into another one
    """
    from itertools import islice

    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(df1, index=True, header=True):
        ws.append(r)

    data = ws.values
    cols = next(data)[1:]

    data = list(data)
    idx = [r[0] for r in data]
    data = (islice(r, 1, None) for r in data)

    df2 = DataFrame(data, index=idx, columns=cols)
    ws = wb.create_sheet()

    for r in dataframe_to_rows(df2, index=True, header=True):
        ws.append(r)

    wb.save("read-write.xlsx")



def using_pandas(df):
    df.to_excel('pandas.xlsx', sheet_name='Sheet1', engine='openpyxl')


def using_xlsxwriter(df):
    df.to_excel('pandas.xlsx', sheet_name='Sheet1', engine='xlsxwriter')


if __name__ == "__main__":
    #df = DataFrame(random.rand(500000, 100))
    df = DataFrame(random.rand(1000, 100))


    #start = time.clock()
    #using_pandas(df)
    #print("pandas openpyxl {0:0.2f}s".format(time.clock()-start))

    #start = time.clock()
    #using_xlsxwriter(df)
    #print("pandas xlsxwriter {0:0.2f}s".format(time.clock()-start))

    start = time.clock()
    openpyxl_in_memory(df)
    print("openpyxl in memory {0:0.2f}s".format(time.clock()-start))

    start = time.clock()
    openpyxl_stream(df)
    print("openpyxl streaming {0:0.2f}s".format(time.clock()-start))


