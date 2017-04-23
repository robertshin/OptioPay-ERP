from __future__ import absolute_import

from pandas import read_sql_query
from pandas.core.frame import DataFrame

from sqlalchemy import create_engine

from with_pandas import dataframe_to_rows

engine = create_engine('postgresql+psycopg2://postgres@localhost:5432/http')

data = read_sql_query("""SELECT cdn, "labelDate", sites FROM cdns WHERE
slice = 'Top100' ORDER by "labelDate" DESC""", engine)

pivoted = data.pivot(index="labelDate", columns="cdn", values="sites")
pivoted.sort_index(ascending=False, inplace=True)

table = dataframe_to_rows(pivoted)

from openpyxl import Workbook
wb = Workbook(write_only=True)
ws = wb.create_sheet("CDNs")

for row in table:
    date = row[0]
    if date:
        row[0] = date.date()
    ws.append(row)

ws = wb.create_sheet("SA")
query = """SELECT cdn_pivot('cdn', '2014-10-01', '2016-02-01', 'Top100')"
"""

conn = engine.raw_connection()
cursor = conn.cursor()
cursor.callproc("cdn_pivot", ['cdn', '2014-10-01', '2016-02-01', 'Top100'])

result = conn.cursor('cdn')
r = result.fetchone()
ws.append([None] + [col.name for col in result.description[1:]])
ws.append(r)

for r in result:
    ws.append(list(r))

wb.save("cdns-24.xlsx")
