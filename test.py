# coding=utf-8

from pexcel.workbook import Workbook
from pexcel.xml_ import (props_app_template, props_core_template, content_types_template,
                         rels_template, workbook_template, workbook_rels_template,
                         worksheet_template, styles_template)



wb = Workbook()
data = [[1, 2, 3], [4, 5, 6]]
ws = wb.new_sheet('hello', data=data)
wb.save('hello.xlsx')
