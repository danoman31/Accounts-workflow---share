"""Fill the X tab and the financial year-end dates on the C tab."""
import argparse
from calendar import monthrange
from datetime import date, datetime
import posixpath
import re
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zoneinfo import ZoneInfo

from working_papers_link import REL, replace_workbook_parts


def append_style_entries(raw, container, entries, size):
    if not entries:
        return raw
    name = container.encode()
    pattern = rb'(<'+ name + rb'\b[^>]*\bcount=")[0-9]+("[^>]*>)(.*?)(</' + name + rb'>)'
    raw, count = re.subn(
        pattern, lambda m: m[1] + str(size).encode() + m[2] + m[3] + b''.join(entries) + m[4],
        raw, flags=re.S,
    )
    if count != 1:
        raise ValueError(f'Could not safely update workbook {container}.')
    return raw


def date_formats(raw, style_ids):
    root = ET.fromstring(raw)
    formats = root.find('{*}numFmts')
    entries = list(formats) if formats is not None else []
    fmt = next((item.get('numFmtId') for item in entries if item.get('formatCode') == 'dd/mm/yyyy'), None)
    if fmt is None:
        fmt = str(max([163, *(int(item.get('numFmtId')) for item in entries)]) + 1)
        entry = f'<numFmt numFmtId="{fmt}" formatCode="dd/mm/yyyy"/>'.encode()
        if formats is None:
            raw = re.sub(rb'(<styleSheet\b[^>]*>)', rb'\1<numFmts count="1">' + entry + b'</numFmts>',
                         raw, count=1)
        else:
            raw = append_style_entries(raw, 'numFmts', [entry], len(entries) + 1)

    styles = list(root.find('{*}cellXfs'))
    result = {}
    new_styles = []
    for index in dict.fromkeys(style_ids):
        old = styles[index]
        attributes = dict(old.attrib, numFmtId=fmt, applyNumberFormat='1')
        match = next((i for i, item in enumerate(styles)
                      if item.attrib == attributes and [ET.tostring(child) for child in item]
                      == [ET.tostring(child) for child in old]), None)
        if match is None:
            item = ET.fromstring(ET.tostring(old))
            item.attrib.update(attributes)
            match = len(styles)
            styles.append(item)
            new_styles.append(ET.tostring(item, encoding='utf-8'))
        result[index] = match
    return append_style_entries(raw, 'cellXfs', new_styles, len(styles)), result


def fill_header(workbook, company, year_end, client_code=None, preparer_initials=None, prepared=None):
    if prepared is None:
        prepared = datetime.now(ZoneInfo('Europe/London')).date()

    def build_parts(source):
        workbook_xml = source.read('xl/workbook.xml')
        root = ET.fromstring(workbook_xml)
        relationships = ET.fromstring(source.read('xl/_rels/workbook.xml.rels'))
        paths = {}
        for tab in ('X', 'C'):
            sheet, = [item for item in root.findall('.//{*}sheet') if item.get('name', '').casefold() == tab.casefold()]
            target, = [item.get('Target') for item in relationships if item.get('Id') == sheet.get(f'{{{REL}}}id')]
            paths[tab] = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
        workbook_properties = root.find('{*}workbookPr')
        uses_1904 = workbook_properties is not None and workbook_properties.get('date1904') in ('1', 'true')
        epoch = date(1904, 1, 1) if uses_1904 else date(1899, 12, 30)

        cells = {'D8': company, 'D10': (year_end - epoch).days, 'G10': year_end.year,
                 'G11': (prepared - epoch).days}
        if client_code is not None:
            cells['D9'] = client_code
        if preparer_initials is not None:
            cells['D11'] = preparer_initials

        tax_dates = {}
        month_end = year_end.day == monthrange(year_end.year, year_end.month)[1]
        for offset in range(3):
            year = year_end.year - offset
            day = monthrange(year, year_end.month)[1] if month_end else year_end.day
            tax_dates[f'G{13 + offset}'] = (date(year, year_end.month, day) - epoch).days

        inputs = {'X': cells, 'C': tax_dates}
        date_cells = {'X': {'D10', 'G11'}, 'C': set(tax_dates)}
        sheets = {tab: source.read(path) for tab, path in paths.items()}
        styles = {}
        for tab, values in inputs.items():
            worksheet = ET.fromstring(sheets[tab])
            for ref in values:
                node = worksheet.find(f'.//{{*}}c[@r="{ref}"]')
                if node is None or node.find('{*}f') is not None:
                    raise ValueError(f'Could not verify working-papers input cell {tab}!{ref}.')
                styles[tab, ref] = int(node.get('s', '0'))
        style_xml, date_style = date_formats(
            source.read('xl/styles.xml'),
            [style for (tab, ref), style in styles.items() if ref in date_cells[tab]],
        )
        for tab, values in inputs.items():
            raw = sheets[tab]
            for ref, value in values.items():
                style = styles[tab, ref]
                if ref in date_cells[tab]:
                    style = date_style[style]
                if isinstance(value, str):
                    body = f't="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is>'
                else:
                    body = f'><v>{value}</v>'
                cell = f'<c r="{ref}" s="{style}" {body}</c>'.encode()
                pattern = rb'<c\b(?=[^>]*\br="' + ref.encode() + rb'")[^>]*?(?:/>|>.*?</c>)'
                raw, count = re.subn(pattern, lambda _: cell, raw, flags=re.S)
                if count != 1:
                    raise ValueError(f'Could not safely update {tab}!{ref}.')
            ET.fromstring(raw)
            sheets[tab] = raw
        workbook_xml = re.sub(
            rb'(<calcPr\b[^>]*?)(/>)',
            lambda m: re.sub(rb'\s+fullCalcOnLoad="[^"]*"', b'', m[1]) + b' fullCalcOnLoad="1"' + m[2],
            workbook_xml,
        )
        ET.fromstring(style_xml)
        return {**{paths[tab]: raw for tab, raw in sheets.items()},
                'xl/styles.xml': style_xml, 'xl/workbook.xml': workbook_xml}

    replace_workbook_parts(workbook, build_parts)
    print(f'Updated X tab and C!G13:G15 in {workbook.name}.', flush=True)


def main():
    from download_xero_reports import DEMO_NAME
    from report_filing import CLIENT_FOLDER
    from working_papers_link import latest_working_papers

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--date', type=date.fromisoformat, required=True, help='Financial year-end (YYYY-MM-DD)')
    parser.add_argument('--client-code', help='Client code, including any leading zeros')
    parser.add_argument('--preparer-initials', help='Preparer initials')
    args = parser.parse_args()
    try:
        fill_header(latest_working_papers(CLIENT_FOLDER), DEMO_NAME, args.date,
                    client_code=args.client_code, preparer_initials=args.preparer_initials)
    except (OSError, ValueError) as error:
        parser.exit(1, f'Could not update the working-papers header: {error}\n')


if __name__ == '__main__':
    main()
