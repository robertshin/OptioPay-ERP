from __future__ import absolute_import

from pandas import read_sql_query, Timestamp
from pandas.core.frame import DataFrame

from sqlalchemy import create_engine
from sqlalchemy import (
    Column,
    Date,
    String,
)
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import (
    relationship,
    backref,
    scoped_session,
    sessionmaker,
)

from openpyxl import Workbook
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.chart.axis import DateAxis, DisplayUnitsLabelList
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart import LineChart, Series, Reference
from openpyxl.drawing.colors import ColorChoice
from openpyxl.worksheet.views import SheetView
from openpyxl.worksheet.page import PageMargins
from openpyxl.chartsheet.properties import ChartsheetProperties

from openpyxl.utils.dataframe import dataframe_to_rows

engine = create_engine('postgresql+psycopg2://postgres@localhost:5432/http')

boundaries = {'renderStart':{'max':10000, 'axis':"Time in seconds", 'title': "Page starts rendering"},
              'onLoad':{'max':20000, 'units':4, 'axis':"Time in seconds", 'title': "Page has loaded"},
              'SpeedIndex':{'max':6000, 'axis':"SpeedIndex\n(smaller is better)", 'title': "SpeedIndex"},
              'PageSpeed':{'min':75, 'max':100, 'axis':"PageSpeed\n(100 is max)", 'title': "PageSpeed"},
              'bytesTotal': {'max':5000000, 'axis':"Total Data in MB", 'title': "Data transferred"}
              }

import time
start = time.time()


query = """
SELECT "labelDate",
"renderStart",
"onLoad",
"SpeedIndex",
"bytesTotal",
label FROM pages
JOIN sites ON
sites.id = pages.id_site
JOIN slices_sites ON
slices_sites.site = sites.id
JOIN browsers on
(browsers.id_browser = pages.id_browser)
WHERE
slice = 'Competition'
AND
"labelDate" >= (SELECT max("labelDate") FROM pages) - interval '1 year 1 month'
AND
"labelDate" < '2015-10-15'
AND browser = 'IE'

UNION
SELECT "labelDate",
"renderStart",
"onLoad",
"SpeedIndex",
"bytesTotal",
label FROM pages
JOIN sites ON
sites.id = pages.id_site
JOIN slices_sites ON
slices_sites.site = sites.id
JOIN browsers on
(browsers.id_browser = pages.id_browser)
WHERE
slice = 'Competition'
AND
"labelDate" >= '2015-10-15'
AND
browser = 'Chrome'
ORDER BY "labelDate", label
"""

data = read_sql_query(query, engine)
pivoted = data.pivot(index="label", columns='labelDate')

def compare():
    q = """
    SELECT "labelDate", "onLoad", browser FROM pages
    JOIN sites on
    (sites.id = pages.id_site)
    JOIN browsers on
    (browsers.id_browser = pages.id_browser)
    WHERE sites.url = 'http://www.bayer.com/'
    AND
    "labelDate" >= (SELECT max("labelDate") FROM pages) - interval '1 year 1 month'
    ORDER BY "labelDate" ASC, browser DESC
    """
    data = read_sql_query(q, engine)
    pivoted = data.pivot(index="browser", columns="labelDate")
    return pivoted

wb = Workbook()
ws = wb.active
wb.remove_sheet(ws)

for sheet in ('renderStart', 'onLoad', 'SpeedIndex', 'bytesTotal',):

    ws = wb.create_sheet(title="{0} data".format(sheet))
    ws.sheet_state = "hidden"
    cs = wb.create_chartsheet(sheet)
    dim = ColumnDimension(ws)
    dim.width = 20.7109375
    ws.column_dimensions['A'] = dim
    chart = LineChart()

    chart.x_axis = DateAxis(crossAx=100)
    chart.y_axis.crossAx = 500
    chart.y_axis.scaling.min = boundaries[sheet].get('min', 0)
    chart.y_axis.scaling.max = boundaries[sheet]['max']
    if sheet in ('renderStart', 'onLoad'):
        chart.y_axis.dispUnits = DisplayUnitsLabelList(builtInUnit="thousands")
    elif sheet == "bytesTotal":
        chart.y_axis.dispUnits = DisplayUnitsLabelList(builtInUnit="millions")
    chart.x_axis.number_format = 'd-mmm-yy'
    chart.x_axis.majorTimeUnit = "months"
    chart.y_axis.title = boundaries[sheet]['axis']
    chart.title = "Relative Performance Bayer, Competitor and Top International Websites: {0}\n Source: httparchive.org - California USA".format(boundaries[sheet]['title'])

    values = pivoted[sheet]
    rows = dataframe_to_rows(values)

    dates = next(rows)
    for col_idx, col in enumerate(dates[1:], 2):
        cell = ws.cell(row=1, column=col_idx, value=col)
        cell.number_format = 'd-mmm-yy'

    for row in rows:
        ws.append(row)

    data = Reference(ws, min_col=1, min_row=2, max_col=len(row), max_row=ws.max_row)
    chart.add_data(data, from_rows=True, titles_from_data=True,)
    dates = Reference(ws, min_col=2, min_row=1, max_col=len(row))
    chart.set_categories(dates)

    bay_idx = 4
    bayer = chart.series[bay_idx]
    bayer.graphicalProperties.line.prstDash = "sysDot"
    bayer.graphicalProperties.line.width = 50050
    #bayer.graphicalProperties.ln.solidFill = "FF0000"
    cs.add_chart(chart)
    cs.sheetViews.sheetView[0].zoomToFit = True
    #cs.sheetViews.sheetView[0].zoomScale = 128


ws = wb.create_sheet("comparison_data")
ws.sheet_state = "hidden"
cs = wb.create_chartsheet("Browser comparison")

table = compare()['onLoad']
rows = dataframe_to_rows(table)
dates = next(rows)
for col_idx, col in enumerate(dates[1:], 2):
    cell = ws.cell(row=1, column=col_idx, value=col)
    cell.number_format = 'd-mmm-yy'

for row in rows:
    ws.append(row)

data = Reference(ws, min_col=1, min_row=2, max_col=len(row), max_row=ws.max_row)
chart = LineChart()

chart.x_axis = DateAxis(crossAx=100)
chart.y_axis.crossAx = 500
chart.y_axis.dispUnits = DisplayUnitsLabelList(builtInUnit="thousands")
chart.x_axis.number_format = 'd-mmm-yy'
chart.x_axis.majorTimeUnit = "months"
chart.add_data(data, from_rows=True, titles_from_data=True,)

chart.title = "Relative Page Load by Browser for bayer.com \n Source: httparchive.org - California USA"

dates = Reference(ws, min_col=2, min_row=1, max_col=len(row))
chart.set_categories(dates)
cs.add_chart(chart)

date = cell.value.strftime("%Y-%b")
wb.save("{0}.xlsx".format(date))
print(time.time() - start)
