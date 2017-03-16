# coding=utf-8

'''
样式部分文档 https://blogs.msdn.microsoft.com/brian_jones/2007/05/29/simple-spreadsheetml-file-part-3-formatting/
'''

from .font import Font


class Style(object):

    def __init__(self):
        self.font = Font()  

    def __eq__(self, other):
        return self.font == other.font

    def __hash__(self):
        return hash(self.font)

    def __str__(self):
        return '{}'.format(self.font)

    def __repr__(self):
        return '<{}>'.format(self.__str__())
