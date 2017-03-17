# coding=utf-8

import os
import sys
path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(path)
from pexcel import Workbook, Worksheet
from pexcel.cell import Cell, Font
from pexcel.template import (props_app_template, props_core_template, content_types_template,
                         rels_template, workbook_template, workbook_rels_template,
                         worksheet_template, styles_template)

wb = Workbook()
data = [['hello', 'world', 'excel'], ['hello', 5.33, 6.78]]
ws = wb.new_sheet('H5', data)
ws['A1'].font = Font(bold=True)
ws['A2'].font = Font(italic=True)
ws['B1'].font = Font(italic=True)

wb.save('hello.xlsx')
