from openpyxl import Workbook
from openpyxl import load_workbook

wb = load_workbook(filename='input.xlsx', read_only=True)

for sheet in wb.sheetnames:

    dest_filename = '{0}.xlsx'.format(sheet)
    new_wb = Workbook()
    del new_wb["Sheet"]

    ws1 = wb[sheet]
    ws2 = new_wb.create_sheet(sheet)

    for row in ws1:
        ws2.append([c.value for c in row])
        first = row[0]
        if first.data_type == "s" and "Total" in first.value:
            for idx in range(len(row)):
                cell = ws2.cell(row=ws2.max_row, column=idx+1)
                bolded = cell.font.copy(bold=True)
                cell.font = bolded

    new_wb.save(dest_filename)
    print("saving {0}".format(dest_filename))

print('finished')