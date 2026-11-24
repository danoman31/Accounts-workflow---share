"""Client folder paths, report versioning, and working-paper hyperlinks."""
import re

from project_paths import OUTPUTS_FOLDER
from working_papers_link import (
    DOCUMENT_LINKS, REPORT_LINKS, STATUTORY_LINKS, WORKING_PAPER_LINKS, WorkbookOpenError,
    latest_working_papers, link_report,
)


CLIENT_FOLDER = OUTPUTS_FOLDER / 'ABC Ltd'
DESTINATION = CLIENT_FOLDER / 'N - Nominal'
ACCOUNTS_FOLDER = CLIENT_FOLDER / 'A - Accounts, PFP_s and Completion'
VAT_FOLDER = CLIENT_FOLDER / 'O - VAT'
DEBTORS_FOLDER = CLIENT_FOLDER / 'H - Debtors'
CREDITORS_FOLDER = CLIENT_FOLDER / 'J - Creditors'
STATUTORY_FOLDER = CLIENT_FOLDER / 'L - Statutory'


def finish_saved_report(code, paths):
    if code in REPORT_LINKS:
        try:
            link_saved_report(code, paths)
        except WorkbookOpenError:
            for path in paths:
                path.unlink()
            raise
    yield from paths


def link_saved_report(code, paths):
    pdf = next(path for path in paths if path.suffix == '.pdf')
    link_working_paper(code, pdf)


def link_working_paper(code, path):
    workbook = latest_working_papers(CLIENT_FOLDER)
    link_report(workbook, code, path)
    tab, cell, _ = WORKING_PAPER_LINKS[code]
    print(f'Updated {workbook.name} {tab}!{cell} → {path.name}', flush=True)


def link_existing_reports():
    folders = {'A': ACCOUNTS_FOLDER, 'N': DESTINATION, 'L1': STATUTORY_FOLDER,
               'H3': DEBTORS_FOLDER, 'J3': CREDITORS_FOLDER}
    for code, (tab, _, _) in {**REPORT_LINKS, **STATUTORY_LINKS}.items():
        folder = folders[tab]
        pattern = re.compile(rf'{re.escape(code)}(?: v([0-9]+))?\.pdf')
        matches = []
        for path in folder.glob('*.pdf'):
            match = pattern.fullmatch(path.name)
            if match and path.is_file():
                matches.append((int(match[1] or '1'), path))
        if matches:
            _, pdf = max(matches)
            link_saved_report(code, (pdf,))
        else:
            print(f'No PDF found for {code}.', flush=True)
    link_working_paper_documents()


def link_working_paper_documents():
    folders = {
        'A': ACCOUNTS_FOLDER,
        'B': CLIENT_FOLDER / 'B - Planning',
        'C': CLIENT_FOLDER / 'C - Taxation',
        'R': CLIENT_FOLDER / 'R - Permanent',
    }
    for code, (tab, _, _) in DOCUMENT_LINKS.items():
        prefix = re.escape(code).replace(r',\ ', r',\s*')
        pattern = re.compile(rf'^{prefix}(?=$|[ _]|-(?![0-9]))', re.I)
        matches = []
        for path in folders[tab].glob('*'):
            if (path.is_file() and path.suffix.lower() in {'.pdf', '.xls', '.xlsx', '.docx', '.doc'}
                    and pattern.match(path.stem)):
                versions = re.findall(r'(?<![A-Za-z0-9])v([0-9]+)(?![A-Za-z0-9])', path.stem, re.I)
                matches.append((int(versions[-1]) if versions else -1, path))
        if matches:
            latest = max(version for version, _ in matches)
            files = [path for version, path in matches if version == latest]
            if len(files) != 1:
                print(f'Multiple files match {code} at the latest version. Link left unchanged.', flush=True)
                continue
            link_working_paper(code, files[0])
