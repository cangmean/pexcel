# coding=utf-8

import os
import sys
path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(path)

from pexcel.workbook import Workbook
from pexcel.worksheet import Worksheet
from pexcel.cell import Cell, Font

data = [['hello', 'world'], ['python', 'language']]
wb = Workbook()
ws = wb.new_sheet("sheet name", data=data)
ws['A1'].font = Font(bold=True)
ws['A2'].font = Font(italic=True)
ws['B1'].font = Font(italic=True)
ws.get_styles()
