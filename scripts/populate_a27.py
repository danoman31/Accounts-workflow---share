"""Copy the Xero Trial Balance nominal codes and year-to-date balances into A27."""
import argparse
from copy import deepcopy
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
from tempfile import TemporaryDirectory
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from report_filing import ACCOUNTS_FOLDER, DESTINATION
from working_papers_link import WorkbookOpenError, require_closed_workbook


FIRST_ROW = 13
SHEET = 'xl/worksheets/sheet1.xml'
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
CELL_REF = re.compile(r'(?<![A-Za-z0-9_])((?:\$?)[A-Z]{1,3}\$?)([1-9][0-9]*)(?![A-Za-z0-9_])')


def latest_file(folder, pattern, label):
    matches = []
    for path in folder.glob('*.xlsx'):
        match = pattern.fullmatch(path.stem)
        if match and path.is_file():
            matches.append((int(match[1] or '0'), path))
    if not matches:
        raise ValueError(f'No {label} workbook found in {folder}.')
    highest = max(version for version, _ in matches)
    paths = [path for version, path in matches if version == highest]
    if len(paths) != 1:
        raise ValueError(f'Multiple {label} workbooks share the latest version.')
    return paths[0]


def adjustment_workbook():
    return latest_file(ACCOUNTS_FOLDER, re.compile(r'A27(?:\s+-\s+Adjustments)?(?:\s+v([0-9]+))?', re.I), 'A27')


def require_closed_a27(workbook):
    try:
        require_closed_workbook(workbook)
    except WorkbookOpenError as error:
        raise WorkbookOpenError(f'Close {workbook.name} in Excel before importing the Trial Balance.') from error


def shared_strings(book):
    if 'xl/sharedStrings.xml' not in book.namelist():
        return []
    return [
        ''.join(node.text or '' for node in entry.findall('.//{*}t'))
        for entry in ET.fromstring(book.read('xl/sharedStrings.xml'))
    ]


def cell_value(cell, strings):
    if cell is None:
        return None
    value = cell.findtext('{*}v')
    if cell.get('t') == 's' and value is not None:
        return strings[int(value)]
    if cell.get('t') == 'inlineStr':
        return ''.join(node.text or '' for node in cell.findall('.//{*}t'))
    return value


def row_cells(row):
    return {re.match(r'[A-Z]+', cell.get('r'))[0]: cell for cell in row.findall('{*}c')}


def amount(value):
    if value in (None, ''):
        return None
    result = Decimal(value)
    if not result.is_finite():
        raise ValueError('The Trial Balance contains a non-finite amount.')
    return result


def trial_balance_accounts(path):
    with ZipFile(path) as book:
        strings = shared_strings(book)
        root = ET.fromstring(book.read(SHEET))
    headers = {
        'account code': 'code', 'account': 'name',
        'debit - year to date': 'debit', 'credit - year to date': 'credit',
    }
    columns = None
    accounts = {}
    for row in root.findall('.//{*}row'):
        values = {column: cell_value(cell, strings) for column, cell in row_cells(row).items()}
        if columns is None:
            found = {headers[value.strip().casefold()]: column for column, value in values.items()
                     if value and value.strip().casefold() in headers}
            if len(found) == len(headers):
                columns = found
            continue
        code = (values.get(columns['code']) or '').strip()
        if code.casefold() == 'total':
            break
        name = (values.get(columns['name']) or '').strip()
        if not code or not name:
            continue
        if code in accounts:
            raise ValueError(f'Trial Balance contains duplicate nominal code {code}.')
        accounts[code] = (name, amount(values.get(columns['debit'])), amount(values.get(columns['credit'])))
    if not accounts:
        raise ValueError('No nominal accounts found in the original Xero Trial Balance.')
    debit = sum((record[1] or 0) for record in accounts.values())
    credit = sum((record[2] or 0) for record in accounts.values())
    if (debit - credit).quantize(Decimal('0.01')):
        raise ValueError('The Trial Balance debit and credit totals do not agree.')
    return accounts


def filled(cell, strings):
    return cell is not None and (cell.find('{*}f') is not None or cell_value(cell, strings) not in (None, ''))


def set_value(row, column, value):
    cells = row_cells(row)
    cell = cells.get(column)
    if cell is None:
        cell = ET.Element(f'{{{NS}}}c', r=f'{column}{row.get("r")}')
        position = next((i for i, name in enumerate(cells) if (len(name), name) > (len(column), column)), len(cells))
        row.insert(position, cell)
    for child in list(cell):
        cell.remove(child)
    if isinstance(value, Decimal):
        cell.attrib.pop('t', None)
        ET.SubElement(cell, f'{{{NS}}}v').text = format(value, 'f')
    else:
        cell.set('t', 'inlineStr')
        inline = ET.SubElement(cell, f'{{{NS}}}is')
        ET.SubElement(inline, f'{{{NS}}}t').text = value


def rewrite_refs(text, threshold, count):
    def shift(match):
        column, number = match.groups()
        row = int(number)
        return f'{column}{row + count if row >= threshold else row}'

    # Double-quoted Excel strings can contain text that looks like a cell reference.
    return '"'.join(part if index % 2 else CELL_REF.sub(shift, part)
                    for index, part in enumerate(text.split('"')))


def add_rows(data, total_row, count):
    rows = {int(row.get('r')): row for row in data.findall('{*}row')}
    sample = rows[FIRST_ROW]
    shared = {}
    for column in ('K', 'L', 'M'):
        formula = row_cells(sample)[column].find('{*}f')
        if formula is not None and formula.get('t') == 'shared':
            shared[column] = formula.get('si')
    for number, row in rows.items():
        if number >= total_row:
            row.set('r', str(number + count))
        for cell in row.findall('{*}c'):
            cell.set('r', rewrite_refs(cell.get('r'), total_row, count))
            formula = cell.find('{*}f')
            if formula is None:
                continue
            if formula.text:
                formula.text = rewrite_refs(formula.text, total_row, count)
                if number == total_row:
                    expected_end = total_row - 1
                    formula.text = re.sub(
                        rf'(:\$?[A-Z]+\$?){expected_end}(?=\))',
                        lambda match: f'{match[1]}{expected_end + count}', formula.text,
                    )
            if formula.get('ref'):
                threshold = total_row - 1 if formula.get('si') in shared.values() else total_row
                formula.set('ref', rewrite_refs(formula.get('ref'), threshold, count))

    insert_at = list(data).index(rows[total_row])
    for offset in range(count):
        number = total_row + offset
        row = deepcopy(sample)
        row.set('r', str(number))
        for cell in row.findall('{*}c'):
            column = re.match(r'[A-Z]+', cell.get('r'))[0]
            cell.set('r', f'{column}{number}')
            cell.attrib.pop('t', None)
            for child in list(cell):
                cell.remove(child)
        for column, formula in {
            'K': f'IF($M{number}>0,$M{number},"")',
            'L': f'IF($M{number}<0,-$M{number},"")',
            'M': f'G{number}-H{number}-E{number}+F{number}+I{number}-J{number}',
        }.items():
            node = ET.SubElement(row_cells(row)[column], f'{{{NS}}}f')
            if column in shared:
                node.attrib.update(t='shared', si=shared[column])
            else:
                node.text = formula
        data.insert(insert_at + offset, row)


def populate_rows(data, strings, accounts):
    rows = {int(row.get('r')): row for row in data.findall('{*}row')}
    total_rows = [
        number for number, row in rows.items() if number >= FIRST_ROW
        and (cell := row_cells(row).get('E')) is not None
        and re.fullmatch(rf'SUM\(E[0-9]+:E{number - 1}\)', cell.findtext('{*}f') or '', re.I)
    ]
    if len(total_rows) != 1:
        raise ValueError('Could not identify the A27 totals row.')
    total_row = total_rows[0]
    existing = {}
    available = []
    for number in range(FIRST_ROW, total_row):
        row = rows.get(number)
        if row is None:
            row = ET.Element(f'{{{NS}}}row', r=str(number))
            following = next((index for index, element in enumerate(data)
                              if int(element.get('r')) > number), len(data))
            data.insert(following, row)
        cells = row_cells(row)
        code = cell_value(cells.get('A'), strings)
        if code:
            if code.strip() in existing:
                raise ValueError(f'A27 already contains duplicate nominal code {code}.')
            existing[code.strip()] = row
        elif all(not filled(cell, strings) or (column in 'KLM' and cell.find('{*}f') is not None)
                 for column, cell in cells.items()):
            available.append(row)

    new_count = len(accounts.keys() - existing.keys())
    added = max(0, new_count - len(available))
    if added:
        add_rows(data, total_row, added)
        available.extend(row for row in data.findall('{*}row')
                         if total_row <= int(row.get('r')) < total_row + added)

    conflicts = []
    changed = 0
    new_rows = iter(available)
    for code, (name, debit, credit) in accounts.items():
        row = existing.get(code)
        if row is None:
            row = next(new_rows)
        inputs = {'A': code, 'B': name, 'E': debit, 'F': credit}
        cells = row_cells(row)
        for column, expected in inputs.items():
            cell = cells.get(column)
            if not filled(cell, strings):
                continue
            current = cell_value(cell, strings)
            if column in ('E', 'F'):
                try:
                    different = (amount(current) or 0) != (expected or 0)
                except InvalidOperation:
                    different = True
            else:
                different = current != expected
            if cell.find('{*}f') is not None or different:
                conflicts.append(code)
                break
        else:
            pending = {column: value for column, value in inputs.items()
                       if value is not None and not filled(cells.get(column), strings)}
            for column, value in pending.items():
                set_value(row, column, value)
            changed += bool(pending)
    return changed, conflicts, total_row, added


def a27_parts(book, accounts):
    raw = book.read(SHEET)
    sheet = ET.fromstring(raw)
    strings = shared_strings(book)
    if [cell_value(sheet.find(f'.//{{*}}c[@r="{ref}"]'), strings) for ref in ('E8', 'F8')] != ['Dr', 'Cr']:
        raise ValueError('A27 must have debit and credit columns at E and F.')
    changed, conflicts, total_row, added = populate_rows(sheet.find('{*}sheetData'), strings, accounts)
    if not changed:
        return {}, set(), changed, conflicts, added
    fragment = ET.tostring(sheet.find('{*}sheetData'), encoding='utf-8')
    raw, matches = re.subn(rb'<sheetData\b[^>]*>.*?</sheetData>', lambda _: fragment, raw, flags=re.S)
    if matches != 1:
        raise ValueError('Could not safely replace the A27 cell data.')
    parts = {}
    omitted = set()
    if added:
        for attribute in (b'ref', b'sqref', b'activeCell', b'topLeftCell'):
            pattern = rb'(<' + rb'(?:dimension|selection|pane|mergeCell)' + rb'\b[^>]*?\b' + attribute + rb'=")([^"]+)(")'
            raw = re.sub(pattern, lambda m: m[1] + rewrite_refs(m[2].decode(), total_row, added).encode() + m[3], raw)
        workbook = book.read('xl/workbook.xml')
        workbook = re.sub(rb'(<definedName\b[^>]*>)([^<]+)(</definedName>)',
                          lambda m: m[1] + rewrite_refs(m[2].decode(), total_row, added).encode() + m[3], workbook)
        parts['xl/workbook.xml'] = workbook
        if 'xl/calcChain.xml' in book.namelist():
            omitted.add('xl/calcChain.xml')
            for part, element, attribute, value in [
                ('xl/_rels/workbook.xml.rels', 'Relationship', 'Target', 'calcChain.xml'),
                ('[Content_Types].xml', 'Override', 'PartName', '/xl/calcChain.xml'),
            ]:
                root = ET.fromstring(book.read(part))
                for child in root.findall(f'{{*}}{element}'):
                    if child.get(attribute) == value:
                        root.remove(child)
                parts[part] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    workbook = parts.get('xl/workbook.xml', book.read('xl/workbook.xml'))
    workbook = re.sub(rb'(<calcPr\b[^>]*?)(/>)',
                      lambda m: re.sub(rb'\s+fullCalcOnLoad="[^"]*"', b'', m[1]) + b' fullCalcOnLoad="1"' + m[2],
                      workbook)
    parts.update({SHEET: raw, 'xl/workbook.xml': workbook})
    ET.fromstring(raw)
    return parts, omitted, changed, conflicts, added


def populate_a27(source, target):
    require_closed_a27(target)
    original = target.stat()
    accounts = trial_balance_accounts(source)
    with TemporaryDirectory(dir=target.parent) as directory:
        staged = Path(directory) / target.name
        with ZipFile(target) as workbook:
            parts, omitted, changed, conflicts, added = a27_parts(workbook, accounts)
            if parts:
                with ZipFile(staged, 'w') as updated:
                    for member in workbook.infolist():
                        if member.filename not in omitted:
                            contents = parts.get(member.filename)
                            updated.writestr(member, workbook.read(member) if contents is None else contents)
        if parts:
            require_closed_a27(target)
            current = target.stat()
            if (current.st_ino, current.st_size, current.st_mtime_ns) != (
                original.st_ino, original.st_size, original.st_mtime_ns
            ):
                raise ValueError('A27 changed during the import. Run the import again.')
            staged.chmod(original.st_mode & 0o777)
            staged.replace(target)
    print(f'A27: updated {changed} nominal rows from {source.name}; added {added} rows.', flush=True)
    if conflicts:
        print(f'A27: existing entries retained for {", ".join(conflicts)}.', flush=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', type=Path, help='Original Xero Trial Balance workbook')
    parser.add_argument('--target', type=Path, help='A27 Adjustments workbook')
    args = parser.parse_args()
    try:
        source = args.source or latest_file(DESTINATION, re.compile(r'N9(?:\s+v([0-9]+))?', re.I), 'N9')
        populate_a27(source, args.target or adjustment_workbook())
    except (ValueError, OSError, InvalidOperation) as error:
        parser.exit(1, f'A27 import failed: {error}\n')


if __name__ == '__main__':
    main()
