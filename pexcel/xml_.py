# coding=utf-8

from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom import minidom
from datetime import datetime
from workbook import Workbook

def prettify(elem):
    """返回一个格式化的xml."""
    rough_string = tostring(elem)
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent='  ', encoding='utf-8')


def rels_template():
    nsmap = {
        'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
    }
    elem_list = [
        ('rId3', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'docProps/app.xml'),
        ('rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'),
        ('rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'xl/workbook.xml')
    ]
    root = Element('Relationships', nsmap)
    for elem in elem_list:
        _id, _type, _target = elem
        kw = {'Id': _id, 'Type': _type, 'Target': _target}
        SubElement(root, 'Relationship', kw)
    return prettify(root)

def content_types_template():
    elem_list = [
        ('Default', 'rels', 'application/vnd.openxmlformats-package.relationships+xml'),
        ('Default', 'xml', 'application/xml'),
        ('Override', '/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'),
        ('Override', '/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'),
        ('Override', '/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
        ('Override', '/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'),
        ('Override', '/xl/worksheets/sheet{}.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
    ]
    nsmap = {
        'xmlns': 'http://schemas.openxmlformats.org/package/2006/content-types'
    }
    root = Element('Types', nsmap)
    for elem in elem_list:
        tag, rels, content_type = elem
        kw = {'contentType': content_type}
        if tag == 'Default':
            kw['Extension'] = rels
        else:
            kw['PartName'] = rels
        SubElement(root, tag, kw)
    return prettify(root)


def props_core_template():
    nsmap = {
        'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
        'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
        'xmlns:dcterms': 'http://purl.org/dc/terms/',
        'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
        'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    }
    root = Element('cp:coreProperties', nsmap)
    SubElement(root, 'dc:creator').text = '{{date}}'
    SubElement(root, 'cp:lastModifiedBy').text = 'pexcel哈哈'.decode('utf-8')
    SubElement(root, 'dcterms:created ', {'xsi:type': 'dcterms:W3CDTF'}).text = datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
    SubElement(root, 'dcterms:modified ', {'xsi:type': 'dcterms:W3CDTF'}).text = datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
    return prettify(root)


def props_app_template(workbook):
    nsmap = {
        'xmlns': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
        'xmlns:vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    }
    root = Element('Properties', nsmap)
    SubElement(root, 'Application').text = 'Microsoft Excel'
    SubElement(root, 'DocSecurity').text = '0'
    SubElement(root, 'ScaleCrop').text = 'false'
    SubElement(root, 'LinksUpToDate').text = 'false'
    SubElement(root, 'SharedDoc').text = 'false'
    SubElement(root, 'HyperlinksChanged').text = 'false'
    SubElement(root, 'AppVersion').text = '14.0300'
    SubElement(root, 'HeadingPairs')
    # heading pairs
    head = SubElement(root, 'HeadingPairs')
    vector = SubElement(head, 'vt:vector', {'size': '2', 'baseType': 'variant'})
    variant = SubElement(vector, 'vt:variant')
    SubElement(variant, 'vt:lpstr').text = 'Worksheets'
    variant = SubElement(vector, 'vt:variant')
    SubElement(variant, 'vt:i4').text = str(workbook.__len__())
    # titles pairs
    titles = SubElement(root, 'TitlesOfParts')
    vector = SubElement(titles, 'vt:vector', {'size': str(workbook.__len__()), 'baseType': 'lpstr'})
    for idx, worksheet in workbook.get_xml_data():
        SubElement(vector, 'vt:lpstr').text = worksheet.name

    return prettify(root)


def workbook_rels_template(workbook):
    nsmap = {
        'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
    }
    root = Element('Relationships', nsmap)
    for idx, sheet in workbook.get_xml_data():
        SubElement(root, 'Relationship', {
            'Id': 'rId{}'.format(idx), 'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
            'Target': 'worksheets/sheet{}.xml'.format(idx)
        })
    return prettify(root)


def workbook_template(workbook):
    nsmap = {
        'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    root = Element('workbook', nsmap)
    SubElement(root, 'fileVersion', {
        'appName': 'xl', 'lastEdited': '5', 'lowestEdited': '5', 'rupBuild': '9303',
    })
    SubElement(root, 'workbookPr', {
        'defaultThemeVersion': '124226',
    })
    book_views = SubElement(root, 'bookViews')
    SubElement(book_views, 'workbookView', {
        'xWindow': '240', 'yWindow': '60', 'windowWidth': '20115', 'windowHeight': '7755',
    })
    sheets = SubElement(root, 'sheets')
    for idx, sheet in workbook.get_xml_data():
        SubElement(sheets, 'sheet', {
            'name': sheet.name.decode('utf-8'), 'sheetId': str(idx), 'r:Id': 'rId{}'.format(idx),
        })
    SubElement(root, 'calcPr', {'calcId': '145621'})
    return prettify(root)


def worksheet_template(worksheet):
    nsmap = {
        'mc:Ignorable': 'x14ac',
        'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'xmlns:x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
    }
    root = Element('worksheet', nsmap)
    sheet_views = SubElement(root, 'sheetViews')
    sheet_view = SubElement(sheet_views, 'sheetView', {'workbookViewId': '0'})
    SubElement(sheet_view, 'selection', {'activeCell': 'A1', 'sqref': 'A1'})
    sheet_data = SubElement(root, 'sheetData')

    return prettify(root)

if __name__ == '__main__':
    wb = Workbook()
    wb.new_sheet('hello')
    wb.new_sheet('谢谢')
    # x = rels_template()
    # x = content_types_template()
    # x = props_core_template()
    x = workbook_rels_template(wb)
    print(x)
