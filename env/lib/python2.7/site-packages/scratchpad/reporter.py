from openpyxl import Workbook
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.chart.axis import DateAxis, DisplayUnitsLabelList
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart import LineChart, Series, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
from pandas import read_sql_query

from sqlalchemy import (
    create_engine,
    func,
    Column,
    Date,
    String,
    Integer,
    Interval,
    PrimaryKeyConstraint,
    union,
    text,
)
from sqlalchemy.ext.declarative import declarative_base, DeferredReflection
from sqlalchemy.orm import sessionmaker

Base = declarative_base(cls=DeferredReflection)

class Page(Base):

    __tablename__ = 'pages'


class Site(Base):

    __tablename__ = 'sites'


class Slice(Base):

    __tablename__ = 'slices'


class SlicesSites(Base):

    __tablename__ = 'slices_sites'
    slice = Column(String(), primary_key=True)
    site = Column(Integer(), primary_key=True)


class Browser(Base):

    __tablename__ = 'browsers'


def query_slice(slice="Competition"):

    engine = create_engine('postgresql+psycopg2://postgres@localhost:5432/http')
    Base.prepare(engine)
    maker = sessionmaker(bind=engine)
    session = maker()

    sq = session.query(func.max(Page.labelDate) - text("interval '1 year 1 month'"))

    q1 = session.query(
        Page.labelDate.label("labelDate"),
        Page.renderStart.label("renderStart"),
        Page.onLoad.label("onLoad"),
        Page.SpeedIndex.label("SpeedIndex"),
        Page.bytesTotal.label("bytesTotal"),
        Site.label.label("label"),
    )
    q1 = q1.join(Site, Site.id==Page.id_site)
    q1 = q1.join(SlicesSites, SlicesSites.site==Site.id)
    q1 = q1.join(Browser, Browser.id_browser==Page.id_browser)
    q1 = q1.filter(SlicesSites.slice==slice)
    q1 = q1.filter(Page.labelDate >= sq.as_scalar())
    q2 = q1.filter(Browser.browser=="Chrome")

    q1 = q1.filter(Page.labelDate<'2015-10-15')
    q1 = q1.filter(Browser.browser=='IE')

    q3 = q1.union(q2)
    q3 = q3.order_by(Page.labelDate, Site.label)

    df = read_sql_query(q3.statement, q3.session.bind)
    return df.pivot(index="label", columns='labelDate')


def query_mobile(slice="Competition"):
    engine = create_engine('postgresql+psycopg2://postgres@localhost:5432/http')
    Base.prepare(engine)
    maker = sessionmaker(bind=engine)
    session = maker()

    sq = session.query(func.max(Page.labelDate) - text("interval '1 year 1 month'"))

    q1 = session.query(
        Page.labelDate.label("labelDate"),
        Page.onLoad.label("onLoad"),
        Site.label.label("label"),
    )
    q1 = q1.join(Site, Site.id==Page.id_site)
    q1 = q1.join(SlicesSites, SlicesSites.site==Site.id)
    q1 = q1.join(Browser, Browser.id_browser==Page.id_browser)
    q1 = q1.filter(SlicesSites.slice==slice)
    q1 = q1.filter(Page.labelDate >= sq.as_scalar())
    q1 = q1.filter(Browser.browser=="Android")

    q1 = q1.order_by(Page.labelDate, Site.label)

    df = read_sql_query(q1.statement, q1.session.bind)
    table = df.pivot(index="label", columns='labelDate')
    return table


wb = Workbook()
wb.remove_sheet(wb.active)

pivoted = query_slice("Bayer")

reports = [query_slice("Competition"), query_mobile('Competition'), query_slice("Bayer"), ]
titles = ['Competition - Desktop', 'Competition - Mobile', 'Bayer Country - Desktop']

for idx, title in enumerate(titles):

    table = reports[idx]
    ws = wb.create_sheet(title="{0} data".format(title))
    ws.sheet_state = "hidden"
    dim = ColumnDimension(ws)
    dim.width = 20.7109375
    ws.column_dimensions['A'] = dim

    values = table['onLoad']
    rows = dataframe_to_rows(values)
    dates = next(rows)
    for col_idx, col in enumerate(dates[1:], 2):
        cell = ws.cell(row=1, column=col_idx, value=col)
        cell.number_format = 'd-mmm-yy'

    for idx, row in enumerate(rows):
        ws.append(row)
        if row[0] == "Bayer":
            bay_idx = idx

    cs = wb.create_chartsheet(title)
    chart = LineChart()
    chart.x_axis = DateAxis(crossAx=100)
    chart.y_axis.crossAx = 500
    chart.x_axis.number_format = 'd-mmm-yy'
    chart.x_axis.majorTimeUnit = "months"
    chart.title = "Relative Performance {0}\n Source: httparchive.org - California USA".format(title)
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 20000
    chart.y_axis.dispUnits = DisplayUnitsLabelList(builtInUnit="thousands")
    chart.y_axis.title = "Time in seconds"

    data = Reference(ws, min_col=1, min_row=2, max_col=len(row), max_row=ws.max_row)
    chart.add_data(data, from_rows=True, titles_from_data=True,)
    dates = Reference(ws, min_col=2, min_row=1, max_col=len(row))
    chart.set_categories(dates)

    bayer = chart.series[bay_idx]
    bayer.graphicalProperties.line.prstDash = "sysDot"
    bayer.graphicalProperties.line.width = 50050
    ##bayer.graphicalProperties.ln.solidFill = "FF0000"
    cs.add_chart(chart)
    cs.sheetViews.sheetView[0].zoomToFit = True

wb.save("{:%Y-%b}.xlsx".format(cell.value))
