"""Read aged report grand totals and update the lead schedules."""
from datetime import date
from decimal import Decimal
import re
from xml.etree import ElementTree as ET

from analyse_nominals import set_input, worksheet_part
from working_papers_link import replace_workbook_parts, report_sheet


AGED_CODES = ('H4', 'J4')


def read_aged_total(page):
    header = page.locator('table tr').filter(has=page.get_by_text('Contact', exact=True))
    header.first.wait_for()
    total_rows = page.locator('table tr').filter(has=page.get_by_text('Total', exact=True))
    amounts = []
    for row in total_rows.all():
        cells = row.locator('th, td').all_inner_texts()
        if cells and cells[0].strip() == 'Total':
            amounts.append(cells[-1].strip())
    if len(amounts) != 1:
        raise ValueError('Could not identify one grand total on the Xero aged summary.')
    value = amounts[0].replace(',', '').replace('£', '').strip()
    if value in ('-', '–', '—'):
        return Decimal(0)
    if value.startswith('(') and value.endswith(')'):
        value = '-' + value[1:-1]
    if not re.fullmatch(r'-?[0-9]+(?:\.[0-9]+)?', value):
        raise ValueError('The Xero aged summary grand total was not a number.')
    return Decimal(value)


def fill_aged_totals(workbook, totals, year_end):
    def build_parts(source):
        workbook_xml = source.read('xl/workbook.xml')
        properties = ET.fromstring(workbook_xml).find('{*}workbookPr')
        uses_1904 = properties is not None and properties.get('date1904') in ('1', 'true')
        epoch = date(1904, 1, 1) if uses_1904 else date(1899, 12, 30)
        x = ET.fromstring(source.read(worksheet_part(source, 'X')))
        if (x.findtext('.//{*}c[@r="D10"]/{*}v') != str((year_end - epoch).days)
                or x.findtext('.//{*}c[@r="G10"]/{*}v') != str(year_end.year)):
            raise ValueError('Fill the X tab for this year-end before writing the aged totals.')
        parts = {}
        for code in AGED_CODES:
            part, raw, sheet = report_sheet(source, code)
            tab = 'H3' if code == 'H4' else 'J3'
            cell = sheet.find('.//{*}c[@r="H9"]')
            if cell is None or sheet.findtext('.//{*}c[@r="H7"]/{*}f') != 'x!G10&"(£)"':
                raise ValueError(f'Could not verify the current-year amount cell {tab}!H9.')
            parts[part] = set_input(raw, cell, totals[code], tab=tab)
            ET.fromstring(parts[part])
        parts['xl/workbook.xml'] = re.sub(
            rb'(<calcPr\b[^>]*?)(/>)',
            lambda m: re.sub(rb'\s+fullCalcOnLoad="[^"]*"', b'', m[1]) + b' fullCalcOnLoad="1"' + m[2],
            workbook_xml,
        )
        return parts

    replace_workbook_parts(workbook, build_parts)
    print(f'Aged totals: H3!H9 = {totals["H4"]:,.2f}; J3!H9 = {totals["J4"]:,.2f}.', flush=True)
