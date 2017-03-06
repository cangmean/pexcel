# coding=utf-8

import re
from .cell import Cell


class Worksheet(object):

    def __init__(self, name, data=None):
        self.name = name
        self.cells = {}
        self.columns = 0
        self.rows = 0
        if data is not None:
            for row_num, row in enumerate(data, 1):
                for col_num, val in enumerate(row, 1):
                    if row_num not in self.cells:
                        self.cells[row_num] = {}
                    self.cells[row_num][col_num] = Cell(row_num, col_num, value=val)
                    self.columns = max(self.columns, col_num)
            else:
                self.rows = len(self.cells.keys())

    def __getitem__(self, key):
        p = re.compile(r'([a-zA-Z]+?)(\d+?)')
        col, row = p.match(key).groups()
        col, row = ord(col.upper()) - 64, int(row)
        if row > self.rows or col > self.columns:
            print(row, col)
            raise IndexError('Key {} 超出 Cell的最大索引。'.format(key))
        return self.cells[row][col]

    def __repr__(self):
        return '<Worksheet {}>'.format(self.name)
