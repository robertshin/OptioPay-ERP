from openpyxl import Workbook
from openpyxl.chart import PieChart, Series, Reference

data = [
    ('Bob', 5),
    ('Alice', 4),
    ('Eve', 2),
]

wb = Workbook()
ws = wb.active
ws.title = "Data"
ws.sheet_state = "hidden"

for r in data:
    ws.append(r)

pie = PieChart()
labels = Reference(ws, min_col=1, max_col=1, min_row=1, max_row=3)
data = Reference(ws, min_col=2, min_row=1, max_row=3)

pie.add_data(data)
pie.set_categories(labels)

cs = wb.create_chartsheet("Chart")
#cs.sheetViews.sheetView[0].zoomToFit = True

cs.add_chart(pie)
wb.save("pie.xlsx")
