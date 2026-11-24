"""Open Xero, set report dates, and capture original downloads."""
from datetime import datetime
from pathlib import Path
import re
from zipfile import is_zipfile

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, expect

from report_filing import CLIENT_FOLDER
from working_papers_link import latest_working_papers, require_closed_workbook


PROFILE = Path.home() / 'Library' / 'Application Support' / 'UHY Xero Reports' / 'Chrome'


def open_xero(playwright):
    PROFILE.mkdir(parents=True, exist_ok=True, mode=0o700)
    context = playwright.chromium.launch_persistent_context(
        PROFILE, channel='chrome', headless=False, chromium_sandbox=True, accept_downloads=True,
    )
    context.pages[0].goto('https://go.xero.com/', wait_until='domcontentloaded')
    return context


def select_demo_company(page):
    switcher = page.get_by_role('button', name=re.compile(r'^Toggle Organisation menu\b'))
    active = page.get_by_role('button', name='Toggle Organisation menu Demo Company (UK)', exact=True)
    try:
        switcher.wait_for(timeout=15000)
        if not active.is_visible():
            switcher.click()
            page.get_by_text('Demo Company (UK)', exact=True).filter(visible=True).click(timeout=8000)
        active.wait_for(timeout=15000)
    except PlaywrightTimeoutError as error:
        raise ValueError(
            'Could not select Demo Company (UK). Run .venv/bin/python scripts/download_xero_reports.py --sign-in '
            'and confirm that Demo Company (UK) is available in the organisation menu.'
        ) from error


def set_report_dates(page, dates):
    update = page.get_by_role('button', name='Update', exact=True)
    update.wait_for()
    fields = page.locator('input:visible:is([type="text"], :not([type])):not([placeholder="Report title"])')
    date_fields = page.locator('input[data-automationid^="report-settings-custom-date-input-"]:visible')
    if date_fields.count() == len(dates):
        fields = date_fields
    fields.first.wait_for()
    if fields.count() != len(dates):
        raise ValueError('Unexpected Xero report date controls.')
    def read_date(field):
        return datetime.strptime(field.input_value().replace(' Sept ', ' Sep '), '%d %b %Y').date()

    expected = [f'{day.day} {day:%b %Y}' for day in dates]
    def verify_dates(timeout):
        for index, value in enumerate(expected):
            displayed = re.escape(value).replace('Sep\\ ', 'Sept?\\ ')
            expect(fields.nth(index)).to_have_value(re.compile(f'^{displayed}$'), timeout=timeout)

    for attempt in range(3):
        if len(dates) == 2:
            select_last_financial_year(page)

        # Set the end first when moving the entire range forward.
        order = list(range(len(dates)))
        if len(dates) == 2 and dates[0] > read_date(fields.nth(1)):
            order.reverse()
        for index in order:
            field = fields.nth(index)
            if read_date(field) != dates[index]:
                field.fill(expected[index])
                field.press('Tab')
        try:
            verify_dates(5000)
            update.click()
            page.get_by_role('button', name='Export', exact=True).wait_for()
            verify_dates(3000)
        except AssertionError as error:
            if attempt == 2:
                shown = ', '.join(field.input_value() for field in fields.all())
                raise ValueError(
                    f'Xero did not keep the requested report dates ({", ".join(expected)}). '
                    f'It displayed {shown}. No PDF or Excel file was saved for this report.'
                ) from error
        else:
            return


def select_last_financial_year(page):
    dropdown = page.get_by_role('button', name='List of convenience dates', exact=True)
    if not dropdown.is_visible():
        return
    dropdown.click()
    option = page.get_by_text('Last financial year', exact=True).filter(visible=True)
    try:
        option.click(timeout=2500)
    except PlaywrightTimeoutError:
        page.keyboard.press('Escape')


def download_report(page, option, path):
    page.get_by_role('button', name='Export', exact=True).click()
    with page.expect_download() as pending:
        page.get_by_text(option, exact=True).click()
    pending.value.save_as(path)


def save_report(report, folder, stem, destination, *, pdf_only=False):
    pdf = folder / f'{stem}.pdf'
    files = []
    if not pdf_only:
        excel = folder / f'{stem}.xlsx'
        download_report(report, re.compile(r'^(Microsoft )?Excel$', re.I), excel)
        if not is_zipfile(excel):
            raise ValueError('Xero returned an unexpected Excel download format.')
        files.append(excel)
    download_report(report, re.compile(r'^PDF$', re.I), pdf)
    if not pdf.read_bytes().startswith(b'%PDF-'):
        raise ValueError('Xero returned an unexpected PDF download format.')
    files.append(pdf)

    require_closed_workbook(latest_working_papers(CLIENT_FOLDER))
    version = 1
    while True:
        name = stem if version == 1 else f'{stem} v{version}'
        published = []
        try:
            for file in files:
                output = destination / f'{name}{file.suffix}'
                output.hardlink_to(file)
                published.append(output)
        except OSError as error:
            for output in published:
                output.unlink()
            if not isinstance(error, FileExistsError):
                raise
            version += 1
        else:
            return tuple(published)
