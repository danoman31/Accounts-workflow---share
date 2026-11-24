"""Update Xero report links without rewriting unrelated workbook parts."""
from pathlib import Path
import posixpath
import re
from tempfile import TemporaryDirectory
from urllib.parse import quote
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'


REPORT_LINKS = {
    'N8': ('N', 'G13', 'Client Nominals (P&L)'),
    'N8-1': ('N', 'G14', 'Client Nominals (B/S)'),
    'N9': ('N', 'G15', 'Client Trial Balance'),
    'N10-1': ('N', 'G16', 'Client P&L'),
    'N10-2': ('N', 'G17', 'Client Balance Sheet'),
    'A101': ('A', 'H34', 'Top Suppliers Report'),
    'A102': ('A', 'H35', 'Top Customers Report'),
    'H4': ('H3', 'F9', 'Debtors as per aged report'),
    'J4': ('J3', 'F9', 'Creditors as per aged report'),
}

STATUTORY_LINKS = {
    'L2': ('L1', 'H43', 'Companies House Company Details'),
    'L3': ('L1', 'H45', 'Companies House Filing History'),
    'L4': ('L1', 'H47', 'Confirmation statement'),
    'L5-1': ('L1', 'H49', 'Directors'),
    'L5-2': ('L1', 'H50', 'PSC'),
}

DOCUMENT_LINKS = {
    'A1': ('A', 'H9', 'Accounts for signature checklist'),
    'A13, A14': ('A', 'H15', 'Directors Minutes & AGM Minutes'),
    'A21, A21-1': ('A', 'H19', 'UHY Non Audit Accountancy Programme - Completion'),
    'A25-1': ('A', 'H23', 'Quality Control Checklist'),
    'A27': ('A', 'H25', 'Opening Balance Adjustments'),
    'A27-1': ('A', 'H26', "VAT - over's & under's form"),
    'A32': ('A', 'H28', 'Disclosure Checklist'),
    'B10': ('B', 'H9', 'UHY Non Audit Accountancy Programme - Planning'),
    'B21': ('B', 'H12', 'Booking in Sheet'),
    'B61-1': ('B', 'H14', 'Budget'),
    'B70': ('B', 'H15', 'UHY Non Audited Accounts Preparation - Risk Assessment'),
    'C1': ('C', 'H10', 'Tax Lead Schedule - UHY Template'),
    'C7': ('C', 'H16', 'Tax Compliance Checklist'),
    'R1': ('R', 'H9', 'Anti Money Laundering Annual Monitoring Review'),
    'R2': ('R', 'H10', 'UHY Client Information Record'),
    'R2-1': ('R', 'H11', 'ID Verification Form'),
}
WORKING_PAPER_LINKS = {**REPORT_LINKS, **STATUTORY_LINKS, **DOCUMENT_LINKS}

def latest_working_papers(folder):
    version_pattern = re.compile(r'(?<![A-Za-z0-9])v([0-9]+)(?![A-Za-z0-9])', re.I)
    candidates = []
    for path in folder.iterdir():
        if path.suffix.lower() != '.xlsx' or path.name.startswith(('~$', '.')) or not path.is_file():
            continue
        try:
            with ZipFile(path) as workbook:
                report_sheet(workbook, 'N9')
        except (BadZipFile, ET.ParseError, IndexError, KeyError, ValueError):
            continue
        versions = version_pattern.findall(path.stem)
        candidates.append((int(versions[-1]) if versions else -1, path))
    if not candidates:
        raise ValueError(f'No working-papers workbook found in {folder}.')
    highest = max(version for version, _ in candidates)
    latest = [path for version, path in candidates if version == highest]
    if len(latest) != 1:
        label = f'version v{highest}' if highest >= 0 else 'no version number'
        raise ValueError(f'Multiple working-papers workbooks have {label}.')
    return latest[0]


class WorkbookOpenError(ValueError):
    pass


def require_closed_workbook(workbook):
    lock_names = {'~$' + workbook.name, '~$' + workbook.name[2:]}
    if any(workbook.with_name(name).exists() for name in lock_names):
        raise WorkbookOpenError('Close the ABC Ltd working-papers workbook in Excel before running reports.')


def link_report(workbook, code, report):
    allowed = {'.pdf', '.xls', '.xlsx', '.docx', '.doc'} if code in DOCUMENT_LINKS else {'.pdf'}
    if report.suffix.lower() not in allowed:
        raise ValueError(f'Unsupported file type for working-papers link {code}.')
    # Excel accepts a literal comma in relative file links. Keep it as a path character.
    target = quote(report.relative_to(workbook.parent).as_posix(), safe='/,')
    replace_workbook_parts(workbook, lambda source: hyperlink_parts(source, target, code))


def replace_workbook_parts(workbook, build_parts):
    require_closed_workbook(workbook)
    original = workbook.stat()
    with TemporaryDirectory(dir=workbook.parent) as staging:
        updated = Path(staging) / workbook.name
        with ZipFile(workbook) as source:
            parts = build_parts(source)
            with ZipFile(updated, 'w') as output:
                for member in source.infolist():
                    contents = parts.pop(member.filename, None)
                    output.writestr(member, source.read(member) if contents is None else contents)
                for name, contents in parts.items():
                    output.writestr(name, contents)
        current = workbook.stat()
        if (current.st_ino, current.st_mtime_ns, current.st_size) != (
            original.st_ino, original.st_mtime_ns, original.st_size
        ):
            raise ValueError('Working papers changed during the export. Retry the workbook update.')
        updated.chmod(original.st_mode & 0o777)
        require_closed_workbook(workbook)
        updated.replace(workbook)


def report_sheet(source, code):
    tab, link_cell, description = WORKING_PAPER_LINKS[code]
    row = link_cell[1:]
    workbook = ET.fromstring(source.read('xl/workbook.xml'))
    sheet, = [item for item in workbook.findall('.//{*}sheet') if item.get('name') == tab]
    workbook_rels = ET.fromstring(source.read('xl/_rels/workbook.xml.rels'))
    sheet_rel, = [item for item in workbook_rels if item.get('Id') == sheet.get(f'{{{REL}}}id')]
    target = sheet_rel.get('Target')
    sheet_path = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
    raw = source.read(sheet_path)
    root = ET.fromstring(raw)
    strings = []
    if 'xl/sharedStrings.xml' in source.namelist():
        strings = [
            ''.join(text.text or '' for text in entry.findall('.//{*}t'))
            for entry in ET.fromstring(source.read('xl/sharedStrings.xml'))
        ]
    code_column, description_column = ('E', 'B') if tab == 'L1' else ('B', 'C')
    if code in ('H4', 'J4'):
        expected_cells = {link_cell: code, f'B{row}': description}
    else:
        expected_cells = {f'{code_column}{row}': code, f'{description_column}{row}': description,
                          link_cell: 'hyperlink'}
    for ref, expected in expected_cells.items():
        cell = root.find(f'.//{{*}}c[@r="{ref}"]')
        value = None
        if cell is not None:
            value = cell.findtext('{*}v')
            if cell.get('t') == 's' and value is not None:
                value = strings[int(value)]
            elif cell.get('t') == 'inlineStr':
                value = ''.join(text.text or '' for text in cell.findall('.//{*}t'))
        if value is None or value.strip().casefold() != expected.casefold():
            raise ValueError(f'Could not verify the working-papers entry at {tab}!{ref}.')
    return sheet_path, raw, root


def hyperlink_parts(source, target, code):
    sheet_path, raw, root = report_sheet(source, code)
    tab, ref, _ = WORKING_PAPER_LINKS[code]
    xml_ref = ref.encode()
    folder, filename = posixpath.split(sheet_path)
    rel_path = posixpath.join(folder, '_rels', f'{filename}.rels')
    try:
        relations = ET.fromstring(source.read(rel_path))
    except KeyError:
        relations = ET.Element(f'{{{PACKAGE_REL}}}Relationships')

    existing = root.find(f'.//{{*}}hyperlink[@ref="{ref}"]')
    relation_id = existing.get(f'{{{REL}}}id') if existing is not None else None
    relation = next((item for item in relations if item.get('Id') == relation_id), None)
    if relation is None:
        ids = {item.get('Id') for item in relations}
        index = 1
        while f'rId{index}' in ids:
            index += 1
        relation_id = f'rId{index}'
        relation = ET.SubElement(relations, f'{{{PACKAGE_REL}}}Relationship', Id=relation_id)
        link = f'<hyperlink ref="{ref}" r:id="{relation_id}"/>'.encode()
        if existing is not None:
            raw, count = re.subn(rb'<hyperlink\b(?=[^>]*\bref="' + xml_ref + rb'")[^>]*/>', link, raw)
        elif root.find('{*}hyperlinks') is not None:
            raw, count = re.subn(rb'</hyperlinks>', link + b'</hyperlinks>', raw)
        elif root.find('{*}printOptions') is not None or root.find('{*}pageMargins') is not None:
            raw, count = re.subn(rb'(?=<(?:printOptions|pageMargins)\b)',
                                 b'<hyperlinks>' + link + b'</hyperlinks>', raw, count=1)
        else:
            raw, count = re.subn(rb'</sheetData>', b'</sheetData><hyperlinks>' + link + b'</hyperlinks>', raw)
        if count != 1:
            raise ValueError(f'Could not safely insert the hyperlink in {tab}!{ref}.')
    relation.set('Type', f'{REL}/hyperlink')
    relation.set('TargetMode', 'External')
    relation.set('Target', target)
    style_index = hyperlink_style_index(source)
    cell_tag = re.compile(rb'<c\b(?=[^>]*\br="' + xml_ref + rb'")[^>]*>')

    def apply_style(match):
        tag = match[0]
        tag = re.sub(rb'\s+s="[^"]*"', b'', tag)
        return tag[:-1] + f' s="{style_index}">'.encode()

    raw, count = cell_tag.subn(apply_style, raw)
    if count != 1:
        raise ValueError(f'Could not safely apply the Hyperlink style to {tab}!{ref}.')
    return {sheet_path: raw, rel_path: ET.tostring(relations, encoding='utf-8', xml_declaration=True)}


def hyperlink_style_index(source):
    styles = ET.fromstring(source.read('xl/styles.xml'))
    named = styles.find('{*}cellStyles/{*}cellStyle[@builtinId="8"]')
    if named is None:
        raise ValueError('The working papers do not define Excel’s built-in Hyperlink style.')
    style_id = named.get('xfId')
    base = styles.find('{*}cellStyleXfs')[int(style_id)]
    for index, cell_style in enumerate(styles.find('{*}cellXfs')):
        if cell_style.get('xfId') == style_id and cell_style.get('fontId') == base.get('fontId'):
            return index
    raise ValueError('The working papers do not have a cell format for Excel’s built-in Hyperlink style.')
