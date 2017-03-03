# coding=utf-8

from zipfile import ZipFile, ZIP_DEFLATED
from .xml_ import (props_app_template, props_core_template, content_types_template,
                   rels_template, workbook_template, workbook_rels_template,
                   worksheet_template, styles_template)


class Writer(object):

    def __init__(self, workbook):
        self.workbook = workbook

    def save(self, f):
        zf = ZipFile(f, 'w', ZIP_DEFLATED)
        zf.writestr('docProps/app.xml', props_app_template(self.workbook))
        zf.writestr('docProps/core.xml', props_core_template())
        zf.writestr('[Content_Types].xml', content_types_template(self.workbook))
        zf.writestr('_rels/.rels', rels_template())
        zf.writestr('xl/styles.xml', styles_template())
        zf.writestr('xl/workbook.xml', workbook_template(self.workbook))
        zf.writestr('xl/_rels/workbook.xml.rels', workbook_rels_template(self.workbook))
        for idx, sheet in self.workbook.get_xml_data():
            zf.writestr('xl/worksheets/sheet{}.xml'.format(idx), worksheet_template(sheet))
        zf.close()
