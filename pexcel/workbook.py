# coding=utf-8

from .worksheet import Worksheet
from .writer import Writer


class Workbook(object):
    encoding = 'utf-8'

    def __init__(self):
        self.worksheets = []
        self.writer = Writer(self)

    def new_sheet(self, sheet_name, data=None):
        sheet = Worksheet(sheet_name, data)
        self.worksheets.append(sheet)
        return sheet

    def get_xml_data(self):
        for idx, sheet in enumerate(self.worksheets, 1):
            yield idx, sheet

    def __len__(self):
        return len(self.worksheets)

    def save(self, filename):
        with open(filename, 'wb') as fp:
            self.writer.save(fp)
