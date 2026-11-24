"""Summarise the N8 revenue and expense transactions on the N6 working paper."""
import argparse
from copy import deepcopy
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
import posixpath
import re
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile

from populate_a27 import cell_value, latest_file, rewrite_refs, row_cells, shared_strings
from report_filing import CLIENT_FOLDER, DESTINATION
from working_papers_link import REL, latest_working_papers, replace_workbook_parts


RULES = (
    (r'\btraining\b', 'Training'),
    (r'\bproject management\b', 'Project management'),
    (r'\bretainer\b', 'Consulting retainers'),
    (r'\b(?:desktop/network support|it support|network support)\b', 'IT support'),
    (r'\belectricity\b', 'Electricity'),
    (r'\bstationery\b', 'Stationery'),
)
REVENUE_TYPES = {'revenue', 'sales', 'other income', 'income'}


def normalise(value):
    return ' '.join(re.sub(r'[^a-z0-9]+', ' ', (value or '').casefold().replace('&', ' and ')).split())


def read_export(path):
    with ZipFile(path) as source:
        sheet = ET.fromstring(source.read('xl/worksheets/sheet1.xml'))
        strings = shared_strings(source)
    for row in sheet.findall('.//{*}row'):
        yield {key: cell_value(cell, strings) for key, cell in row_cells(row).items()}


def nominal_groups(n8, n9, year_end):
    accounts = {}
    for row in read_export(n9):
        code, name, kind = (row.get(key) for key in ('A', 'B', 'C'))
        if code and name and kind and code != 'Account Code':
            account = normalise(name)
            if account in accounts:
                raise ValueError(f'Duplicate account name in N9: {name}.')
            accounts[account] = (code, normalise(kind) in REVENUE_TYPES)
    if not accounts:
        raise ValueError('No account codes and account types found in N9.')

    groups = defaultdict(lambda: defaultdict(Decimal))
    active = None
    seen_headers = False
    valid_period = False
    wanted_end = f'{year_end.day} {year_end:%B %Y}'
    for row in read_export(n8):
        title = row.get('A') or ''
        if title.startswith('For the period '):
            start_text, separator, end_text = title.removeprefix('For the period ').partition(' to ')
            from download_xero_reports import financial_year_start

            if separator:
                start = datetime.strptime(start_text, '%d %B %Y').date()
                valid_period = start == financial_year_start(year_end) and end_text == wanted_end
        if title == 'Date':
            if any(row.get(key) != value for key, value in [('C', 'Description'), ('E', 'Debit'), ('F', 'Credit')]):
                raise ValueError('Unexpected N8 transaction columns.')
            seen_headers = True
        elif seen_headers and title and not row.get('B'):
            active = accounts.get(normalise(title))
        elif seen_headers and active and row.get('B'):
            code, revenue = active
            debit, credit = (Decimal(row.get(col) or '0') for col in ('E', 'F'))
            net = credit - debit if revenue else debit - credit
            description = (row.get('C') or '').casefold()
            label = next((label for pattern, label in RULES if re.search(pattern, description)), 'Needs review')
            groups[code][label] += net
    if not (seen_headers and valid_period):
        raise ValueError('N8 does not match the requested financial year.')
    return groups


def worksheet_part(source, tab):
    workbook = ET.fromstring(source.read('xl/workbook.xml'))
    sheet, = [sheet for sheet in workbook.findall('.//{*}sheet') if sheet.get('name', '').casefold() == tab.casefold()]
    relations = ET.fromstring(source.read('xl/_rels/workbook.xml.rels'))
    target, = [rel.get('Target') for rel in relations if rel.get('Id') == sheet.get(f'{{{REL}}}id')]
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))


def set_input(raw, cell, value, *, tab="N6"):
    ref = cell.get('r')
    if cell.find('{*}f') is not None:
        raise ValueError(f'{tab}!{ref} contains a formula.')
    if isinstance(value, str):
        body = f'<is><t xml:space="preserve">{escape(value)}</t></is>'
        kind = ' t="inlineStr"'
    else:
        body = f'<v>{value:f}</v>'
        kind = ''
    pattern = rb'(<(?:\w+:)?c\b(?=[^>]*\br="' + ref.encode() + rb'")([^>]*?))(?:/>|>.*?</(?:\w+:)?c>)'

    def replace(match):
        attributes = re.sub(rb'\s+t="[^"]*"', b'', match[1])
        prefix = re.match(rb'<((?:\w+:)?c)\b', attributes)[1].decode()
        namespace = f' xmlns="{cell.tag.split("}")[0][1:]}"' if ':' in prefix else ''
        return attributes + f'{kind}{namespace}>{body}</{prefix}>'.encode()

    raw, count = re.subn(pattern, replace, raw, flags=re.S)
    if count != 1:
        raise ValueError(f'Could not safely update {tab}!{ref}.')
    return raw


def insert_detail_rows(raw, subtotal, count):
    sheet = ET.fromstring(raw)
    data = sheet.find('{*}sheetData')
    rows = {int(row.get('r')): row for row in data.findall('{*}row')}
    sample = deepcopy(rows[subtotal - 1])
    for cell in sample.findall('{*}c'):
        cell.attrib.pop('t', None)
        for child in list(cell):
            cell.remove(child)
    for number, row in rows.items():
        if number >= subtotal:
            row.set('r', str(number + count))
        for cell in row.findall('{*}c'):
            cell.set('r', rewrite_refs(cell.get('r'), subtotal, count))
            formula = cell.find('{*}f')
            if formula is None:
                continue
            if formula.text:
                formula.text = rewrite_refs(formula.text, subtotal, count)
                if number == subtotal:
                    formula.text = re.sub(
                        rf'(:\$?[A-Z]+\$?){subtotal - 1}(?=\))',
                        lambda match: f'{match[1]}{subtotal + count - 1}', formula.text,
                    )
            if formula.get('ref'):
                formula.set('ref', rewrite_refs(formula.get('ref'), subtotal, count))

    position = list(data).index(rows[subtotal])
    for offset in range(count):
        number = subtotal + offset
        row = deepcopy(sample)
        row.set('r', str(number))
        for column, cell in row_cells(row).items():
            cell.set('r', f'{column}{number}')
        data.insert(position + offset, row)
    fragment = ET.tostring(data, encoding='utf-8')
    raw, replaced = re.subn(rb'<(?:\w+:)?sheetData\b[^>]*>.*?</(?:\w+:)?sheetData>',
                            lambda _: fragment, raw, flags=re.S)
    if replaced != 1:
        raise ValueError('Could not safely insert N6 grouping rows.')
    raw = re.sub(rb'(<(?:dimension|selection|pane)\b[^>]*?\b(?:ref|sqref|activeCell|topLeftCell)=")([^"]+)(")',
                 lambda m: m[1] + rewrite_refs(m[2].decode(), subtotal, count).encode() + m[3], raw)
    raw = re.sub(rb'(<brk\b[^>]*?\bid=")([0-9]+)(")',
                 lambda m: m[1] + str(int(m[2]) + (count if int(m[2]) >= subtotal else 0)).encode() + m[3], raw)
    return raw


def shift_n6_metadata(source, book_xml, insertions):
    parts = {}
    if not insertions:
        return parts, book_xml
    workbook = ET.fromstring(book_xml)
    sheet_id = next(sheet.get('sheetId') for sheet in workbook.findall('.//{*}sheet') if sheet.get('name') == 'N6')
    if 'xl/calcChain.xml' in source.namelist():
        chain = ET.fromstring(source.read('xl/calcChain.xml'))
        active = None
        for cell in chain:
            active = cell.get('i', active)
            if active == sheet_id:
                for subtotal, count in insertions:
                    cell.set('r', rewrite_refs(cell.get('r'), subtotal, count))
        parts['xl/calcChain.xml'] = ET.tostring(chain, encoding='utf-8', xml_declaration=True)

    for subtotal, count in insertions:
        book_xml = re.sub(
            rb'(<definedName\b[^>]*>)([^<]*(?:N6|&apos;N6&apos;)[^<]*![^<]+)(</definedName>)',
            lambda m: m[1] + rewrite_refs(m[2].decode(), subtotal, count).encode() + m[3],
            book_xml,
        )
    return parts, book_xml


def analyse_nominals(workbook, n8, n9, year_end):
    groups = nominal_groups(n8, n9, year_end)

    def build_parts(source):
        header = ET.fromstring(source.read(worksheet_part(source, 'X')))
        year = header.find('.//{*}c[@r="G10"]')
        end = header.find('.//{*}c[@r="D10"]')
        book_xml = source.read('xl/workbook.xml')
        props = ET.fromstring(book_xml).find('{*}workbookPr')
        epoch = date(1904, 1, 1) if props is not None and props.get('date1904') in ('1', 'true') else date(1899, 12, 30)
        if (year is None or year.findtext('{*}v') != str(year_end.year)
                or end is None or end.findtext('{*}v') != str((year_end - epoch).days)):
            raise ValueError('Fill the X tab for the requested year before updating N6.')
        part = worksheet_part(source, 'N6')
        raw = source.read(part)
        sheet = ET.fromstring(raw)
        strings = shared_strings(source)
        rows = {int(row.get('r')): row_cells(row) for row in sheet.findall('.//{*}row')}
        insertions = []
        for index, cells in reversed(list(rows.items())):
            code = cell_value(cells.get('A'), strings)
            if not code or not re.fullmatch(r'[0-9]+', code.strip()):
                continue
            code = code.strip()
            values = groups.get(code)
            if not values:
                print(f'N6 {code}: no matching N8 movement.', flush=True)
                continue
            start = index + 1
            end = None
            for number in range(start, max(rows) + 1):
                total = rows.get(number, {}).get('H')
                if total is not None and total.findtext('{*}f') == f'SUM(H{start}:H{number - 1})':
                    end = number
                    break
            if end is None:
                raise ValueError(f'Could not find the N6 subtotal for {code}.')
            slots = {}
            empty = []
            # Reserve the first detail row beneath the nominal heading.
            for number in range(start + 1, end):
                name, value = (rows[number].get(key) for key in ('B', 'H'))
                if name is None or value is None:
                    raise ValueError(f'N6 row {number} does not have both input cells.')
                label = cell_value(name, strings)
                if label:
                    if label not in values or label in slots:
                        raise ValueError(f'Existing N6 analysis needs review before updating nominal {code}.')
                    slots[label] = (name, value)
                elif cell_value(value, strings) in (None, ''):
                    empty.append((name, value))
                else:
                    raise ValueError(f'Existing N6 value has no grouping name on row {number}.')
            last = end - 1
            last_used = any(int(value.get('r')[1:]) == last for _, value in slots.values())
            usable = [slot for slot in empty if int(slot[1].get('r')[1:]) != last]
            extra = max(len(values) - len(slots) - len(usable) + int(last_used), 0)
            if extra:
                raw = insert_detail_rows(raw, end, extra)
                insertions.append((end, extra))
                current = ET.fromstring(raw)
                inserted = {int(row.get('r')): row_cells(row) for row in current.findall('.//{*}row')
                            if end <= int(row.get('r')) < end + extra - 1}
                empty.extend((cells['B'], cells['H']) for cells in inserted.values())
            else:
                empty = usable
            for label, amount in values.items():
                name, value = slots[label] if label in slots else empty.pop(0)
                raw = set_input(raw, name, label)
                raw = set_input(raw, value, amount)
            print(f'N6 {code}: {len(values)} group(s), total {sum(values.values()):,.2f}.', flush=True)
        ET.fromstring(raw)
        parts, book_xml = shift_n6_metadata(source, book_xml, insertions)
        book_xml = re.sub(
            rb'(<calcPr\b[^>]*?)(/>)',
            lambda m: re.sub(rb'\s+fullCalcOnLoad="[^"]*"', b'', m[1]) + b' fullCalcOnLoad="1"' + m[2],
            book_xml,
        )
        return {**parts, part: raw, 'xl/workbook.xml': book_xml}

    replace_workbook_parts(workbook, build_parts)
    print(f'Updated N6 analysis in {workbook.name}.', flush=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--date', type=date.fromisoformat, required=True, help='Financial year-end (YYYY-MM-DD)')
    parser.add_argument('--n8', type=Path, help='Account Transactions workbook')
    parser.add_argument('--n9', type=Path, help='Trial Balance workbook')
    args = parser.parse_args()
    try:
        n8 = args.n8 or latest_file(DESTINATION, re.compile(r'N8(?: v([0-9]+))?', re.I), 'N8')
        n9 = args.n9 or latest_file(DESTINATION, re.compile(r'N9(?: v([0-9]+))?', re.I), 'N9')
        analyse_nominals(latest_working_papers(CLIENT_FOLDER), n8, n9, args.date)
    except (OSError, ValueError, ArithmeticError) as error:
        parser.exit(1, f'Nominal analysis failed: {error}\n')


if __name__ == '__main__':
    main()
