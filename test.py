# coding=utf-8

from pexcel import Workbook, Worksheet
from pexcel.template import (props_app_template, props_core_template, content_types_template,
                         rels_template, workbook_template, workbook_rels_template,
                         worksheet_template, styles_template)


data = [['hello', 'world', 'excel'], ['hello', 5.33, 6.78]]
# ws = Worksheet('H5', data)
# print(ws.cells)
# for row_num, columns in ws.cells.iteritems():
#     for col_num, cell in columns.iteritems():
#         print(cell.value, cell.val_type)

wb = Workbook()
wb.new_sheet('hello', data)
wb.save('hello.xlsx')

x = styles_template()
print(x)
