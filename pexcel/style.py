# coding=utf-8

'''
样式部分文档 https://blogs.msdn.microsoft.com/brian_jones/2007/05/29/simple-spreadsheetml-file-part-3-formatting/
'''

from .font import Font


class Style(object):

    def __init__(self):
        self.font = Font()

    @property
    def is_default(self):
        return self.hash_key == Style().hash_key

    @property
    def hash_key(self):
        return hash((self.font))

    def __eq__(self, other):
        return self.hash_key == other.hash_key

    def __hash__(self):
        return self.hash_key

    def __str__(self):
        return '{}'.format(self.font)

    def __repr__(self):
        return '({})'.format(self.__str__())
