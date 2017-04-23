from pympler.muppy import print_summary
from openpyxl.cell import Cell
d = [Cell(None, 'A', 1) for i in range(100000)]
print_summary()
