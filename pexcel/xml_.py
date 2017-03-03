# coding=utf-8

from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom import minidom
from datetime import datetime
from collections import OrderedDict as order_dict


def prettify(elem):
    """返回一个格式化的xml."""
    rough_string = tostring(elem)
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent='  ', encoding='UTF-8')


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

def content_types_template(workbook):
    elem_list = [
        ('Default', 'rels', 'application/vnd.openxmlformats-package.relationships+xml'),
        ('Default', 'xml', 'application/xml'),
        ('Override', '/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'),
        ('Override', '/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'),
        ('Override', '/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
        #('Override', '/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'),
    ]
    nsmap = {
        'xmlns': 'http://schemas.openxmlformats.org/package/2006/content-types'
    }
    root = Element('Types', nsmap)
    for elem in elem_list:
        tag, rels, content_type = elem
        kw = {'ContentType': content_type}
        if tag == 'Default':
            kw['Extension'] = rels
        else:
            kw['PartName'] = rels
        SubElement(root, tag, kw)
    for idx, sheet in workbook.get_xml_data():
        SubElement(root, 'Override', {
            'PartName': '/xl/worksheets/sheet{}.xml'.format(idx),
            'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
        })
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
    SubElement(root, 'dc:creator').text = 'pexcel'
    SubElement(root, 'cp:lastModifiedBy').text = 'pexcel'
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
    for idx, sheet in workbook.get_xml_data():
        SubElement(vector, 'vt:lpstr').text = sheet.name

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
    # SubElement(root, 'Relationship', {
    #     'Id': 'rId1000', 'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
    #     'Target': 'styles.xml'
    # })
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
            'name': sheet.name.decode('utf-8'), 'sheetId': str(idx), 'r:id': 'rId{}'.format(idx),
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
    sheet_format = SubElement(root, 'sheetFormatPr', {
        'defaultRowHeight': '15', 'x14ac:dyDescent': '0.25'
    })
    sheet_data = SubElement(root, 'sheetData')
    for row_num, columns in worksheet.cells.iteritems():
        row = SubElement(sheet_data, 'row', {'r': str(row_num)})
        for col_num, value in columns.iteritems():
            col = SubElement(row, 'c', {'r': '{}{}'.format(chr(64+col_num), row_num)})
            SubElement(col, 'v').text = str(value).decode('utf-8')
    SubElement(root, 'pageMargins', {
        'left': '0.7', 'right': '0.7', 'top': '0.75',
        'bottom': '0.75', 'header': '0.3', 'footer': '0.3'
    })
    return prettify(root)


def styles_template():
    nsmap = {
        'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'xmlns:x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
        'mc:Ignorable': 'x14ac',
    }
    root = Element('styleSheet', nsmap)
    return prettify(root)
