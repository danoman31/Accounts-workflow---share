"""Run the Account Transactions filters and locate completed UK VAT returns."""
from calendar import monthrange
from datetime import date, timedelta
from pathlib import Path
import re
from tempfile import TemporaryDirectory
from urllib.parse import urlsplit

from report_filing import finish_saved_report
from xero_report_browser import save_report, set_report_dates


def export_account_transactions(report, dates, destination, *, link_reports=True):
    report.get_by_placeholder('Search for accounts').click()
    report.locator('#select-all button').click()
    report.keyboard.press('Escape')
    report.get_by_text(re.compile(r'^[1-9][0-9]* accounts selected$')).filter(visible=True).wait_for()
    filters = (
        ('N8', {'Revenue', 'Expense'}),
        ('N8-1', {'Asset', 'Liability', 'Equity'}),
    )
    for stem, types in filters:
        set_account_types(report, types)
        set_report_dates(report, dates)
        with TemporaryDirectory(dir=destination) as staging:
            paths = save_report(report, Path(staging), stem, destination)
        yield from (finish_saved_report(stem, paths) if link_reports else paths)


def set_account_types(report, selected):
    report.locator('[data-automationid="report-settings-filter-button"]').click()
    trigger = report.locator('[data-automationid="report-settings-filter-accordion-trigger"]')
    account_type = trigger.filter(has=report.get_by_text('Account Type', exact=True))
    expand = account_type.get_by_role('button', name='Expand', exact=True)
    if expand.is_visible():
        expand.click()
    for name in ('Asset', 'Liability', 'Expense', 'Revenue', 'Equity'):
        choice = report.get_by_role('checkbox', name=name, exact=True)
        choice.wait_for(state='attached')
        if choice.is_checked() != (name in selected):
            report.locator('[data-automationid="report-settings-list-filter-checkbox"]').filter(
                has=choice
            ).click()
        if choice.is_checked() != (name in selected):
            raise ValueError(f'Xero did not apply the {name} selection.')
    report.get_by_role('button', name=re.compile(r'^Apply\b')).click()


def vat_quarters(year_end):
    if year_end.day != monthrange(year_end.year, year_end.month)[1]:
        raise ValueError('Use a month-end financial year-end for the VAT quarter schedule.')
    end = year_end
    for number in range(1, 5):
        year, month_index = divmod(end.year * 12 + end.month - 1 - 3, 12)
        month = month_index + 1
        previous_end = date(year, month, monthrange(year, month)[1])
        yield f'O5-{number}', previous_end + timedelta(days=1), end
        end = previous_end


def vat_period_pattern(start, end):
    def displayed(day):
        month = day.strftime('%B')
        words = {month, month[:3]}
        if month == 'September':
            words.add('Sept')
        names = '|'.join(sorted(words, key=len, reverse=True))
        return rf'0?{day.day}\s+(?:{names})\s+{day.year}'

    return re.compile(rf'\b{displayed(start)}\s*(?:to|[-–—])\s*{displayed(end)}\b', re.I)


def export_vat_returns(overview, year_end, destination):
    quarters = list(vat_quarters(year_end))
    completed = overview.get_by_text('Completed', exact=True).filter(visible=True)
    completed.wait_for()
    empty = overview.get_by_text('Finalised VAT returns will be shown here', exact=True)
    # Restrict navigation to return rows following the Completed heading.
    rows = completed.locator('xpath=following::tr | following::*[@role="row"]')
    for stem, start, end in quarters:
        period = vat_period_pattern(start, end)
        matches = rows.filter(has_text=period).filter(visible=True)
        if empty.is_visible() or not matches.count():
            print(f'VAT quarter ending {end:%d %B %Y}: no completed return found ({stem}.pdf).', flush=True)
            continue
        if matches.count() != 1:
            print(f'VAT quarter ending {end:%d %B %Y}: multiple completed returns found ({stem}.pdf).', flush=True)
            continue
        row = matches.first
        # Follow a view link only. Never use Create, Review or filing controls.
        link = row.get_by_role('link', name=period).or_(
            row.get_by_role('link', name=re.compile(r'^View(?: VAT)? return$|^View$', re.I))
        )
        if link.count() != 1:
            print(f'VAT quarter ending {end:%d %B %Y}: no unique view link ({stem}.pdf).', flush=True)
            continue
        url = link.evaluate('(element) => element.href')
        address = urlsplit(url)
        if address.scheme != 'https' or not (address.hostname or '').endswith('.xero.com'):
            raise ValueError('Unexpected Xero VAT return URL.')
        report = overview.context.new_page()
        try:
            report.goto(url, wait_until='domcontentloaded')
            with TemporaryDirectory(dir=destination) as staging:
                yield from save_report(report, Path(staging), stem, destination, pdf_only=True)
        finally:
            report.close()
