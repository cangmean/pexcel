# coding=utf-8

import re
from .cell import Cell


class Worksheet(object):

    def __init__(self, name, data=None):
        self.name = name
        self.cells = {}
        self.cell_list = []
        self.columns = 0
        self.rows = 0
        self.styles = {}
        if data is not None:
            for row_num, row in enumerate(data, 1):
                for col_num, val in enumerate(row, 1):
                    if row_num not in self.cells:
                        self.cells[row_num] = {}
                    cell = Cell(row_num, col_num, value=val)
                    self.cells[row_num][col_num] = cell
                    self.cell_list.append(cell)
                    self.columns = max(self.columns, col_num)
            else:
                self.rows = len(self.cells.keys())

    def get_styles(self):
        styles = {}
        style_list = []
        items = {'fonts': {}}
        for cell in self.cell_list:
            style_list.append(cell.style)
        for index, style in enumerate(style_list):
            styles[style] = styles.get(style, len(styles) + 1)
            setattr(style, 'id', styles[style])
        for style in styles.keys():
            obj = style.font
            items['fonts'][obj] = items['fonts'].get(obj, len(items['fonts']) + 1)
            obj.id = items['fonts'][obj]
        for k, v in items.items():
            items[k] = [tup[0] for tup in sorted(v.items(), key=lambda x: x[1])]
        print(items)

    def __getitem__(self, key):
        p = re.compile(r'([a-zA-Z]+?)(\d+?)')
        col, row = p.match(key).groups()
        col, row = ord(col.upper()) - 64, int(row)
        if row > self.rows or col > self.columns:
            raise IndexError('Key {} 超出 Cell的最大索引.'.format(key))
        return self.cells[row][col]

    def __repr__(self):
        return "<Worksheet '{}'>".format(self.name)
