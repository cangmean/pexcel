# coding=utf-8


class CellError(Exception):
    pass


class CellIndexError(CellError):
    pass


class Cell(object):
    '''
    描述文档: http://officeopenxml.com/SScontentOverview.php
    Excel 中单元格中的t制定类型
    比如: <c r='A1' t='s'> 类型的可能值为:
    b: 布尔类型
    d: 日期类型
    e: 错误类型
    inlinestr: 用于内联字符串（即，不存储在共享字符串部分中，但直接在单元格中）
    n: 数字
    s: 用于共享字符串（因此存储在共享字符串部分而不是单元格中）
    str: 用于公式
    ==========================
    pexcel 采用一下三种类型

    <c r="A1" t="inlineStr">
        <is>
            <t>我的字符串</t>
        </is>
    </c>
    <c r="A1" t="n">
        <v> 400 </v>
    </c>

    <c r="A1" t="str">
        <f> SUM（B2：B8）</f>
        <v> 2105 </v>
    </c>
    '''

    def __init__(self, row, column, value=None):
        self.row = row
        self.column = column
        self.value = value
        self.name = '{}{}'.format(chr(64+self.column), self.row)

    def _get_styles(self):
        pass

    def _set_styles(self, styles):
        pass

    styles = property(_get_styles, _set_styles)

    @property
    def val_type(self):
        value = self.value
        if isinstance(value, str):
            if len(value) > 0 and value[0] == '=':
                return 'FORMULA'
            else:
                return 'STRING'
        elif isinstance(value, (float, int)):
            return 'NUMBER'
        else:
            # 其他类型全部设为string.
            return 'STRING'

    @property
    def excel_format(self):
        if self.val_type == 'FORMULA':
            return 'str'
        elif self.val_type == 'NUMBER':
            return 'n'
        else:
            return 'inlineStr'

    def __repr__(self):
        return '<Cell {}>'.format(self.name)

if __name__ == '__main__':
    a = Cell(row=1, column=2, value=33)
    print(a.val_type)
