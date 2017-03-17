# coding=utf-8

import os
import sys
path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(path)

from pexcel.workbook import Workbook
from pexcel.worksheet import Worksheet
from pexcel.cell import Cell, Font
from pexcel.template import (props_app_template, props_core_template, content_types_template,
                         rels_template, workbook_template, workbook_rels_template,
                         worksheet_template, styles_template)

data = [['hello', 'world'], ['python', 'language']]
wb = Workbook()
ws = wb.new_sheet("sheet name", data=data)
ws['A1'].font = Font(bold=True)
ws['A2'].font = Font(italic=True)
ws['B1'].font = Font(italic=True)

ws.get_styles()

# x = styles_template(wb)
# print(x)
