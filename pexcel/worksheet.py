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
        self._styles = {}
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

    @property
    def styles(self):
        self.get_styles()
        return self._styles

    def get_styles(self):
        styles = {}
        style_list = []
        items = {'fonts': {}}

        for cell in self.cell_list:
            style_list.append(cell.style)

        # style list是所有cell style的集合
        # 通过dict key hash 的特性来去重
        # styles { style: style.id } 并对 原有的style设置id
        for index, style in enumerate(style_list):
            styles[style] = styles.get(style, len(styles) + 1)
            setattr(style, 'id', styles[style])

        for style in styles.keys():
            obj = style.font
            # 用相同的方式对 font 等属性去重并添加id
            items['fonts'][obj] = items['fonts'].get(obj, len(items['fonts']) + 1)
            obj.id = items['fonts'][obj]
        for k, v in items.items():
            items[k] = [tup[0] for tup in sorted(v.items(), key=lambda x: x[1])]
        print(items)
        self._styles = [tup[0] for tup in sorted(styles.items(), key=lambda x: x[1])]
        for i in self._styles:
            print(i.id, i.font.id)
        return items

    def __getitem__(self, key):
        p = re.compile(r'([a-zA-Z]+?)(\d+?)')
        col, row = p.match(key).groups()
        col, row = ord(col.upper()) - 64, int(row)
        if row > self.rows or col > self.columns:
            raise IndexError('Key {} 超出 Cell的最大索引.'.format(key))
        return self.cells[row][col]

    def __repr__(self):
        return "<Worksheet '{}'>".format(self.name)
