from zipfile import ZipFile
from openpyxl.xml.functions import fromstring
from openpyxl.xml.constants import CONTYPES_NS

from openpyxl import load_workbook
#wb = load_workbook("Issues/bug467.xlsm", keep_vba=True)
#wb.save("Issues/bug467-1.xlsm")

archive = ZipFile("comments.xlsx")
files = set(archive.namelist())

ct = archive.read("[Content_Types].xml")
ct = fromstring(ct)
parts = ct.findall('{%s}Override' % CONTYPES_NS)
parts = set(p.attrib["PartName"][1:] for p in parts)

print(parts - files)
