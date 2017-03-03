# coding=utf-8


class Worksheet(object):

    def __init__(self, name, data=None, force_name=False):
        self.name = name
        self.cells = {}
        self.columns = 0
        self.rows = 0
        if data is not None:
            for x, row in enumerate(data, 1):
                for y, cell in enumerate(row, 1):
                    if x not in self.cells:
                        self.cells[x] = {}
                    self.cells[x][y] = cell
                    self.columns = max(self.columns, y)
            else:
                self.rows = len(self.cells.keys())


if __name__ == '__main__':
    data = [[1, 2, 3], [4, 5, 6]]
    ws = Worksheet('H5', data)
    print(ws.cells)
