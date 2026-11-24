"""Save public Companies House records and link them in the ABC Ltd working papers."""
import argparse
from pathlib import Path
import re
from tempfile import TemporaryDirectory
from urllib.parse import urlsplit

from playwright.sync_api import Error as PlaywrightError, sync_playwright

from report_filing import CLIENT_FOLDER, STATUTORY_FOLDER, link_working_paper
from working_papers_link import latest_working_papers, require_closed_workbook


REGISTER = 'https://find-and-update.company-information.service.gov.uk'


def company_number(value):
    number = value.strip().upper()
    if not re.fullmatch(r'[A-Z0-9]{8}', number):
        raise argparse.ArgumentTypeError('Enter the eight-character Companies House number, including leading zeros.')
    return number


def locate_company(page, number):
    url = f'{REGISTER}/company/{number}'
    response = page.goto(url, wait_until='load')
    if response is None or not response.ok or urlsplit(page.url).path.rstrip('/') != f'/company/{number}':
        raise ValueError(f'Could not open Companies House company {number}.')
    record_number = page.locator('#company-number').inner_text().strip().removeprefix('Company number').strip()
    if record_number != number:
        raise ValueError(f'Companies House returned company {record_number} instead of {number}.')
    company = page.get_by_role('heading', level=1).inner_text().strip()
    return url, company


def print_page(page, url, company, output):
    page.goto(url, wait_until='load')
    page.get_by_role('heading', name=company, exact=True).wait_for()
    page.pdf(path=output, format='A4', print_background=True,
             margin={'top': '12mm', 'right': '12mm', 'bottom': '12mm', 'left': '12mm'})


def confirmation_statement(page, output):
    while True:
        filing_type = page.locator('td.filing-type').filter(has_text=re.compile(r'^\s*CS01\s*$'))
        row = page.locator('tr').filter(has=filing_type).first
        if row.count():
            link = row.get_by_role('link', name=re.compile(r'^View PDF', re.I))
            url = link.evaluate('(element) => element.href')
            if urlsplit(url).netloc != urlsplit(REGISTER).netloc:
                raise ValueError('Unexpected Companies House filing link.')
            response = page.request.get(url, timeout=30000)
            if not response.ok:
                raise ValueError(f'Companies House returned HTTP {response.status} for the confirmation statement.')
            output.write_bytes(response.body())
            return row.get_by_role('cell').first.inner_text().strip()
        next_page = page.get_by_role('link', name=re.compile(r'^Next(?: page)?$', re.I))
        if not next_page.count():
            raise ValueError('No confirmation statement was found in Companies House filing history.')
        next_page.click()


def publish_pdf(path, workbook):
    version = 1
    while True:
        require_closed_workbook(workbook)
        suffix = '' if version == 1 else f' v{version}'
        target = STATUTORY_FOLDER / f'{path.stem}{suffix}.pdf'
        try:
            target.hardlink_to(path)
        except FileExistsError:
            version += 1
        else:
            return target


def download_company(number, *, link_reports=True):
    workbook = latest_working_papers(CLIENT_FOLDER)
    require_closed_workbook(workbook)
    with TemporaryDirectory(dir=STATUTORY_FOLDER) as staging, sync_playwright() as playwright:
        folder = Path(staging)
        browser = playwright.chromium.launch(channel='chrome', chromium_sandbox=True)
        try:
            page = browser.new_page()
            url, company = locate_company(page, number)
            print(f'Companies House: {company} ({number})', flush=True)
            print_page(page, url, company, folder / 'L2.pdf')
            print_page(page, url + '/filing-history', company, folder / 'L3.pdf')
            filed = confirmation_statement(page, folder / 'L4.pdf')
            print_page(page, url + '/officers', company, folder / 'L5-1.pdf')
            print_page(page, url + '/persons-with-significant-control', company, folder / 'L5-2.pdf')
        finally:
            browser.close()

        for code in ('L2', 'L3', 'L4', 'L5-1', 'L5-2'):
            if not (folder / f'{code}.pdf').read_bytes().startswith(b'%PDF-'):
                raise ValueError(f'Companies House returned an invalid PDF for {code}.')
        for code in ('L2', 'L3', 'L4', 'L5-1', 'L5-2'):
            saved = publish_pdf(folder / f'{code}.pdf', workbook)
            print(f'Saved Companies House PDF: {saved.name}', flush=True)
            if not link_reports:
                continue
            try:
                link_working_paper(code, saved)
            except (ValueError, OSError):
                saved.unlink()
                raise
        print(f'Latest confirmation statement filed: {filed}', flush=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('company_number', type=company_number, help='Companies House company number')
    parser.add_argument('--download-only', action='store_true', help='Save PDFs without updating workbook links')
    args = parser.parse_args()
    try:
        download_company(args.company_number, link_reports=not args.download_only)
    except (OSError, ValueError, PlaywrightError) as error:
        parser.exit(1, f'Companies House export failed: {error}\n')


if __name__ == '__main__':
    main()
