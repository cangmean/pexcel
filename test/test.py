# coding=utf-8


from pyexcelerate import Workbook, Color, Style, Font, Fill, Format
from datetime import datetime

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws.set_cell_value(1, 1, 1)
ws.set_cell_style(1, 1, Style(font=Font(bold=True)))
ws.set_cell_style(1, 1, Style(font=Font(italic=True)))
ws.set_cell_style(1, 1, Style(font=Font(underline=True)))
ws.set_cell_style(1, 2, Style(font=Font(underline=True)))

wb._align_styles()
print(wb._items)
