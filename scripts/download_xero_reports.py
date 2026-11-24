"""Download original Xero reports or run the legacy API exporter."""
import argparse
import base64
from calendar import monthrange
from datetime import date, timedelta
from decimal import Decimal
import hashlib
from http.server import BaseHTTPRequestHandler, HTTPServer
import json
from pathlib import Path
import re
import secrets
import ssl
import subprocess
import sys
import textwrap
from tempfile import TemporaryDirectory
import time
from urllib.error import HTTPError
from urllib.parse import parse_qs, urlencode, urlsplit
from urllib.request import Request, urlopen
import webbrowser
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from local_credentials import read_credential, write_credential
from project_paths import PROJECT_ROOT
from report_filing import (
    ACCOUNTS_FOLDER, CLIENT_FOLDER, CREDITORS_FOLDER, DEBTORS_FOLDER, DESTINATION, VAT_FOLDER,
    finish_saved_report,
    link_existing_reports,
)
from working_papers_link import WorkbookOpenError, latest_working_papers, require_closed_workbook


TEMPLATE = PROJECT_ROOT / 'trial_balance_template.xlsx'
DESTINATION_FOLDER = DESTINATION
REDIRECT = 'http://localhost:8765/callback'
XERO = 'https://api.xero.com'
DEMO_NAME = 'Demo Company (UK)'
KEYCHAIN_SERVICE = 'Accounting Workflow Xero Trial Balance'
TOKEN_URL = 'https://identity.xero.com/connect/token'
PROFIT_AND_LOSS_SCOPE = 'accounting.reports.profitandloss.read'
BALANCE_SHEET_SCOPE = 'accounting.reports.balancesheet.read'
REPORT_SCOPES = ('accounting.reports.trialbalance.read', PROFIT_AND_LOSS_SCOPE, BALANCE_SHEET_SCOPE)
REQUESTED_SCOPES = ' '.join((*REPORT_SCOPES, 'offline_access'))
SHEET = 'xl/worksheets/sheet1.xml'
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', NS)


def request_json(url, token=None, tenant=None, form=None):
    headers = {'Accept': 'application/json'}
    if token:
        headers['Authorization'] = f'Bearer {token}'
    if tenant:
        headers['xero-tenant-id'] = tenant
    data = None if form is None else urlencode(form).encode()
    try:
        with urlopen(Request(url, data=data, headers=headers), timeout=30,
                     context=xero_tls_context()) as response:
            return json.load(response)
    except HTTPError as error:
        if url == TOKEN_URL and error.code == 400:
            try:
                details = json.load(error)
            except ValueError:
                details = {}
            if details.get('error') == 'invalid_grant':
                raise ExpiredGrant('Xero sign-in has expired or was revoked.') from None
        raise ValueError(f'Xero returned HTTP {error.code} for {urlsplit(url).path}.') from None


class ExpiredGrant(ValueError):
    pass


def load_session():
    saved = read_credential('xero.json')
    if saved is not None:
        return saved
    result = subprocess.run(
        ['/usr/bin/security', 'find-generic-password', '-s', KEYCHAIN_SERVICE,
         '-a', DEMO_NAME, '-w'], capture_output=True, text=True,
    )
    if result.returncode == 44:
        return None
    if result.returncode:
        raise ValueError(f'Could not read macOS Keychain: {result.stderr.strip()}')
    try:
        saved = json.loads(base64.urlsafe_b64decode(result.stdout.strip()))
    except ValueError:
        raise ValueError('The saved Xero session in macOS Keychain is unreadable.') from None
    save_session(saved)
    return saved


def save_session(session):
    write_credential('xero.json', session)


def demo_connection(token):
    connections = request_json(f'{XERO}/connections', token)
    demos = [item for item in connections if item.get('tenantName') == DEMO_NAME]
    if len(demos) != 1:
        raise ValueError(f'Expected one authorised {DEMO_NAME} connection.')
    return demos[0]['tenantId']


def xero_session(*, interactive=True, required_scopes=()):
    saved = load_session()
    needs_consent = not saved or not set(required_scopes).issubset(saved.get('scopes', []))
    if not needs_consent and saved['expires_at'] > time.time() + 60:
        return saved['access_token'], demo_connection(saved['access_token'])

    if needs_consent and not interactive:
        raise ExpiredGrant('Authorise Xero in Terminal before running bot exports.')
    client_id = saved['client_id'] if saved else input(
        'Enter the Client ID from Xero Configuration: '
    ).strip()
    if not client_id:
        raise ValueError('Enter your Xero Client ID.')
    if not needs_consent:
        try:
            response = request_json(TOKEN_URL, form={
                'grant_type': 'refresh_token', 'client_id': client_id,
                'refresh_token': saved['refresh_token'],
            })
        except ExpiredGrant:
            if not interactive:
                raise
            response = authorise(client_id)
            tenant = demo_connection(response['access_token'])
        else:
            tenant = None
    else:
        response = authorise(client_id)
        tenant = demo_connection(response['access_token'])

    token = response['access_token']
    scopes = response.get('scope', '')
    if not scopes:
        scopes = ' '.join(saved.get('scopes', [])) if tenant is None else REQUESTED_SCOPES
    if not set(required_scopes).issubset(scopes.split()):
        raise ExpiredGrant('Xero did not grant the requested report permission.')
    save_session({
        'client_id': client_id, 'access_token': token,
        'refresh_token': response['refresh_token'],
        'expires_at': time.time() + int(response['expires_in']),
        'scopes': scopes.split(),
    })
    return token, tenant or demo_connection(token)


def xero_tls_context():
    context = ssl.create_default_context()
    # The python.org macOS installer can leave its CA bundle unconfigured.
    macos_ca_bundle = Path('/etc/ssl/cert.pem')
    if context.cert_store_stats()['x509_ca'] == 0 and macos_ca_bundle.is_file():
        context.load_verify_locations(cafile=str(macos_ca_bundle))
    return context


def authorise(client_id):
    verifier = secrets.token_urlsafe(48)
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode()).digest()
    ).decode().rstrip('=')
    state = secrets.token_urlsafe(32)
    result = {}

    class Callback(BaseHTTPRequestHandler):
        def log_message(self, *_):
            pass

        def do_GET(self):
            location = urlsplit(self.path)
            params = parse_qs(location.query)
            if location.path != '/callback' or params.get('state') != [state]:
                self.send_error(400, 'Invalid callback')
                return
            result.update(params)
            self.send_response(200)
            self.send_header('Content-Type', 'text/plain; charset=utf-8')
            self.end_headers()
            self.wfile.write(b'Xero sign-in response received. Return to Terminal.')

    query = urlencode({
        'response_type': 'code', 'client_id': client_id, 'redirect_uri': REDIRECT,
        'scope': REQUESTED_SCOPES,
        'state': state,
        'code_challenge': challenge, 'code_challenge_method': 'S256',
    })
    link = 'https://login.xero.com/identity/connect/authorize?' + query
    with HTTPServer(('127.0.0.1', 8765), Callback) as server:
        print(f'Select {DEMO_NAME} on the Xero consent screen. If the browser does not open, visit:\n{link}', flush=True)
        webbrowser.open(link)
        server.timeout = 1
        deadline = time.monotonic() + 300
        while not result and time.monotonic() < deadline:
            server.handle_request()

    if not result.get('code'):
        raise ValueError('Xero sign-in did not complete. Check consent and the registered redirect URI.')
    return request_json(TOKEN_URL, form={
        'grant_type': 'authorization_code', 'client_id': client_id,
        'code': result['code'][0], 'redirect_uri': REDIRECT, 'code_verifier': verifier,
    })


def trial_balance_rows(payload, as_at):
    reports = payload.get('Reports', [])
    if len(reports) != 1 or reports[0].get('ReportID') != 'TrialBalance':
        raise ValueError('Xero did not return a Trial Balance.')
    report = reports[0]
    titles = report['ReportTitles']
    acceptable_dates = {
        f'As at {as_at.day} {as_at:%B %Y}',
        f'As at {as_at:%d %B %Y}',
    }
    if DEMO_NAME not in titles or acceptable_dates.isdisjoint(titles):
        raise ValueError(f'Unexpected report organisation or date: {titles}')

    rows = [[title] for title in titles] + [[]]
    totals = [Decimal(0)] * 4
    summary = None
    account_count = 0
    for section in report['Rows']:
        if section.get('Title'):
            rows.append([section['Title']])
        for row in section.get('Rows', [section]):
            values = [cell['Value'] for cell in row['Cells']]
            kind = row['RowType']
            if kind in ('Row', 'SummaryRow'):
                if len(values) != 5:
                    raise ValueError('Unexpected Trial Balance columns.')
                values[1:] = [Decimal(value) if value else None for value in values[1:]]
            if kind == 'Row':
                account_count += 1
                totals = [total + (value or 0) for total, value in zip(totals, values[1:])]
            elif kind == 'SummaryRow':
                summary = values[1:]
            rows.append(values)
    if not account_count or totals != summary:
        raise ValueError('Account rows did not reconcile to Xero’s summary.')
    return rows


def save_workbook(rows, destination, sheet_name='Trial Balance'):
    # Reuse the original workbook's column widths and number format.
    with ZipFile(TEMPLATE) as template:
        worksheet = ET.fromstring(template.read(SHEET))
        workbook = ET.fromstring(template.read('xl/workbook.xml'))
        workbook.find(f'.//{{{NS}}}sheet').set('name', sheet_name)
        if sheet_name != 'Trial Balance':
            worksheet.find(f'{{{NS}}}cols/{{{NS}}}col[@min="2"]').set('width', '25')
        data = worksheet.find(f'{{{NS}}}sheetData')
        data.clear()
        for number, values in enumerate(rows, 1):
            row = ET.SubElement(data, f'{{{NS}}}row', r=str(number))
            for column, value in zip('ABCDE', values):
                if value is None:
                    continue
                ref = f'{column}{number}'
                if isinstance(value, Decimal):
                    cell = ET.SubElement(row, f'{{{NS}}}c', r=ref, s='1')
                    ET.SubElement(cell, f'{{{NS}}}v').text = str(value)
                else:
                    cell = ET.SubElement(row, f'{{{NS}}}c', r=ref, t='inlineStr')
                    inline = ET.SubElement(cell, f'{{{NS}}}is')
                    ET.SubElement(inline, f'{{{NS}}}t').text = value
        destination.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(destination, 'x', ZIP_DEFLATED) as output:
            for name in template.namelist():
                if name == SHEET:
                    content = ET.tostring(worksheet, encoding='utf-8')
                elif name == 'xl/workbook.xml':
                    content = ET.tostring(workbook, encoding='utf-8')
                else:
                    content = template.read(name)
                output.writestr(name, content)


def save_pdf(rows, destination):
    lines = []
    for row in rows:
        if len(row) < 2:
            lines.append(''.join(row))
            continue
        account, *amounts = row
        labels = textwrap.wrap(account, width=32) or ['']
        lines.extend(labels[:-1])
        values = [f'{value:,.2f}' if isinstance(value, Decimal) else value or ''
                  for value in amounts]
        width = 25 if len(values) == 1 else 13
        lines.append(f'{labels[-1]:32}' + ''.join(f'{value:>{width}}' for value in values))
    with destination.open('xb') as output:
        result = subprocess.run(
            ['/usr/sbin/cupsfilter', '-i', 'text/plain', '-m', 'application/pdf'],
            input=('\n'.join(lines) + '\n').encode(), stdout=output,
            stderr=subprocess.PIPE, timeout=60,
        )
    if result.returncode or destination.stat().st_size == 0:
        raise ValueError('macOS could not generate the report PDF.')


def save_versioned_reports(rows, folder, basename='N9', sheet_name='Trial Balance'):
    with TemporaryDirectory(dir=folder) as staging:
        workbook = Path(staging) / 'N9.xlsx'
        pdf = Path(staging) / 'N9.pdf'
        save_workbook(rows, workbook, sheet_name)
        save_pdf(rows, pdf)
        version = 1
        while True:
            name = basename if version == 1 else f'{basename} v{version}'
            excel_output = folder / f'{name}.xlsx'
            pdf_output = folder / f'{name}.pdf'
            try:
                excel_output.hardlink_to(workbook)
            except FileExistsError:
                version += 1
                continue
            try:
                pdf_output.hardlink_to(pdf)
            except FileExistsError:
                excel_output.unlink()
                version += 1
            except OSError:
                excel_output.unlink()
                raise
            else:
                return excel_output, pdf_output


def check_export_paths():
    if not DESTINATION_FOLDER.is_dir():
        raise ValueError(f'Destination folder is missing at {DESTINATION_FOLDER}')
    if not TEMPLATE.is_file():
        raise ValueError(f'Excel template is missing at {TEMPLATE}')


def export_trial_balance(as_at, *, interactive=True):
    check_export_paths()
    token, tenant = xero_session(interactive=interactive)
    url = f'{XERO}/api.xro/2.0/Reports/TrialBalance?date={as_at.isoformat()}'
    rows = trial_balance_rows(request_json(url, token, tenant), as_at)
    return save_versioned_reports(rows, DESTINATION_FOLDER)


def financial_report_rows(payload, report_id, period_label, max_amount_columns):
    reports = payload.get('Reports', [])
    if len(reports) != 1 or reports[0].get('ReportID') != report_id:
        raise ValueError(f'Xero did not return the expected report ({report_id}).')
    report = reports[0]
    if DEMO_NAME not in report['ReportTitles']:
        raise ValueError('Unexpected report organisation.')
    rows = [[report['ReportName']], [DEMO_NAME], [period_label], []]

    def append_rows(report_rows):
        for row in report_rows:
            if row.get('Title'):
                rows.append([row['Title']])
            if row.get('Cells'):
                label, *values = [cell['Value'] for cell in row['Cells']]
                if len(values) > max_amount_columns:
                    raise ValueError('Xero returned more report columns than expected.')
                if row['RowType'] in ('Row', 'SummaryRow'):
                    values = [Decimal(value) if value else None for value in values]
                rows.append([label, *values])
            append_rows(row.get('Rows', []))

    append_rows(report['Rows'])
    if len(rows) == 4:
        raise ValueError('Xero returned no report rows.')
    return rows


def financial_year_start(year_end):
    prior_year = year_end.year - 1
    prior_month_end = monthrange(prior_year, year_end.month)[1]
    current_month_end = monthrange(year_end.year, year_end.month)[1]
    prior_day = prior_month_end if year_end.day == current_month_end else year_end.day
    prior_year_end = year_end.replace(year=prior_year, day=prior_day)
    return prior_year_end + timedelta(days=1)


def export_profit_and_loss(as_at, *, interactive=True):
    check_export_paths()
    start = financial_year_start(as_at)
    token, tenant = xero_session(interactive=interactive, required_scopes=(PROFIT_AND_LOSS_SCOPE,))
    query = urlencode({'fromDate': start.isoformat(), 'toDate': as_at.isoformat()})
    url = f'{XERO}/api.xro/2.0/Reports/ProfitAndLoss?{query}'
    rows = financial_report_rows(request_json(url, token, tenant), 'ProfitAndLoss',
                                 f'{start} to {as_at}', 1)
    return save_versioned_reports(rows, DESTINATION_FOLDER, 'N10-1', 'Profit and Loss')


def export_balance_sheet(as_at, *, interactive=True):
    if as_at.day != monthrange(as_at.year, as_at.month)[1]:
        raise ValueError('Use a month-end date for the Xero Balance Sheet.')
    check_export_paths()
    token, tenant = xero_session(interactive=interactive, required_scopes=(BALANCE_SHEET_SCOPE,))
    url = f'{XERO}/api.xro/2.0/Reports/BalanceSheet?date={as_at.isoformat()}'
    rows = financial_report_rows(request_json(url, token, tenant), 'BalanceSheet',
                                 f'As at {as_at}', 2)
    return save_versioned_reports(rows, DESTINATION_FOLDER, 'N10-2', 'Balance Sheet')


def api_main(argv=None):
    parser = argparse.ArgumentParser(description='Generate reports from Xero API data.')
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument('--date', type=date.fromisoformat, help='As-at date (YYYY-MM-DD)')
    action.add_argument('--profit-and-loss', type=date.fromisoformat, metavar='DATE',
                        help='Export Profit and Loss for the 12 months ending on DATE')
    action.add_argument('--balance-sheet', type=date.fromisoformat, metavar='DATE',
                        help='Export the Balance Sheet as at month-end DATE')
    action.add_argument('--authorize', action='store_true', help='Connect Xero without exporting a workbook')
    args = parser.parse_args(argv)
    try:
        if args.authorize:
            xero_session(required_scopes=REPORT_SCOPES)
            print(f'Xero connection ready for {DEMO_NAME}.')
            return
        if args.profit_and_loss:
            destinations = export_profit_and_loss(args.profit_and_loss)
        elif args.balance_sheet:
            destinations = export_balance_sheet(args.balance_sheet)
        else:
            destinations = export_trial_balance(args.date)
    except (ValueError, OSError, KeyError, ArithmeticError, EOFError, subprocess.TimeoutExpired) as error:
        parser.exit(1, f'Export failed: {error}\n')
    for destination in destinations:
        print(f'Saved: {destination}')



def run_native_reports(year_end, *, statements=False, all_reports=False,
                       top_suppliers=False, top_customers=False, account_transactions=False,
                       vat_returns=False, aged_summaries=False, link_reports=True, aged_totals=None):
    from playwright.sync_api import expect, sync_playwright
    from fill_aged_totals import AGED_CODES, read_aged_total
    from xero_report_browser import open_xero, save_report, select_demo_company, set_report_dates
    from xero_special_reports import export_account_transactions, export_vat_returns

    require_closed_workbook(latest_working_papers(CLIENT_FOLDER))
    reports = []
    if all_reports or not (statements or top_suppliers or top_customers or account_transactions or vat_returns or aged_summaries):
        reports.append(('Trial Balance', 'N9', (year_end,), DESTINATION, None, False))
    if statements or all_reports:
        reports.extend([
            ('Profit and Loss', 'N10-1', (financial_year_start(year_end), year_end), DESTINATION, None, False),
            ('Balance Sheet', 'N10-2', (year_end,), DESTINATION, None, False),
        ])
    if top_suppliers or all_reports:
        reports.append(('Payable Invoice Detail', 'A101',
                        (financial_year_start(year_end), year_end), ACCOUNTS_FOLDER, 'Top suppliers', True))
    if top_customers or all_reports:
        reports.append(('Receivable Invoice Detail', 'A102',
                        (financial_year_start(year_end), year_end), ACCOUNTS_FOLDER, 'Top customers', True))
    if account_transactions or all_reports:
        reports.append(('Account Transactions', None,
                        (financial_year_start(year_end), year_end), DESTINATION, None, False))
    if aged_summaries or all_reports:
        reports.extend([
            ('Aged Receivables Summary', 'H4', (year_end,), DEBTORS_FOLDER, None, True),
            ('Aged Payables Summary', 'J4', (year_end,), CREDITORS_FOLDER, None, True),
        ])
    if vat_returns or all_reports:
        reports.append(('UK VAT Return', None, (year_end,), VAT_FOLDER, None, True))
    with sync_playwright() as playwright:
        context = open_xero(playwright)
        try:
            dashboard = context.pages[0]
            select_demo_company(dashboard)
            dashboard.get_by_role('button', name='Reporting', exact=True).click()
            dashboard.get_by_role('link', name='All reports', exact=True).click()
            for title, stem, dates, destination, preset, pdf_only in reports:
                link = dashboard.get_by_role('link', name=title, exact=True).first
                url = link.evaluate('(element) => element.href')
                if urlsplit(url).scheme != 'https' or not (urlsplit(url).hostname or '').endswith('.xero.com'):
                    raise ValueError('Unexpected Xero report URL.')
                report = context.new_page()
                try:
                    report.goto(url, wait_until='domcontentloaded')
                    if title == 'Account Transactions':
                        yield from export_account_transactions(report, dates, destination, link_reports=link_reports)
                        continue
                    if title == 'UK VAT Return':
                        yield from export_vat_returns(report, year_end, destination)
                        continue
                    if preset:
                        report.get_by_text(preset, exact=True).filter(visible=True).click()
                        expect(report.get_by_placeholder('Report title')).to_have_value(
                            re.compile(rf'^{re.escape(preset)}$', re.I)
                        )
                    set_report_dates(report, dates)
                    if aged_totals is not None and stem in AGED_CODES:
                        aged_totals[stem] = read_aged_total(report)
                    with TemporaryDirectory(dir=destination) as staging:
                        paths = save_report(report, Path(staging), stem, destination, pdf_only=pdf_only)
                    yield from (finish_saved_report(stem, paths) if link_reports else paths)
                finally:
                    report.close()
        finally:
            context.close()


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if argv[:1] == ['--api']:
        api_main(argv[1:])
        return
    from playwright.sync_api import Error as PlaywrightError, sync_playwright
    from download_companies_house import company_number, download_company
    from xero_report_browser import open_xero

    parser = argparse.ArgumentParser(description=__doc__, epilog='For API-generated reports, run with --api --help.')
    action = parser.add_mutually_exclusive_group()
    action.add_argument('--statements', action='store_true',
                        help='Download Profit and Loss and Balance Sheet instead of Trial Balance')
    action.add_argument('--sign-in', action='store_true', help='Set up the dedicated Xero browser session')
    action.add_argument('--all-reports', action='store_true', help='Download all configured reports')
    action.add_argument('--top-suppliers', action='store_true', help='Download Top Suppliers only')
    action.add_argument('--top-customers', action='store_true', help='Download Top Customers only')
    action.add_argument('--account-transactions', action='store_true',
                        help='Download both Account Transactions reports only')
    action.add_argument('--aged-summaries', action='store_true', help='Download aged receivables and payables PDFs only')
    action.add_argument('--vat-returns', action='store_true', help='Download completed UK VAT return PDFs only')
    action.add_argument('--link-existing', action='store_true', help='Link the latest saved working-paper PDFs without opening Xero')
    parser.add_argument('--date', type=date.fromisoformat, default=date(2026, 3, 31),
                        help='Financial year-end in YYYY-MM-DD format')
    parser.add_argument('--company-number', type=company_number,
                        help='Include Companies House filings for this number with --all-reports')
    parser.add_argument('--client-code', help='Client code for the working-papers X tab')
    parser.add_argument('--preparer-initials', help='Preparer initials for the working-papers X tab')
    parser.add_argument('--download-only', action='store_true',
                        help='Save reports without updating working-paper links')
    args = parser.parse_args(argv)
    if (args.client_code is not None or args.preparer_initials is not None) and (not args.all_reports or args.download_only):
        parser.error('Working-papers details require --all-reports without --download-only.')
    if args.company_number and not args.all_reports:
        parser.error('--company-number requires --all-reports')
    if args.download_only and (args.sign_in or args.link_existing or args.company_number):
        parser.error('--download-only is for Xero exports. Run Companies House separately for download-only mode.')
    if args.link_existing:
        try:
            link_existing_reports()
        except WorkbookOpenError as error:
            parser.exit(3, f'{error}\n')
        return
    if args.sign_in:
        with sync_playwright() as playwright:
            context = open_xero(playwright)
            try:
                input('Sign in to Xero and complete MFA. Press Enter once the Xero dashboard loads. ')
            finally:
                context.close()
        return
    aged_totals = {} if (args.all_reports or args.aged_summaries) and not args.download_only else None
    adjustment = None
    trial_balance = None
    nominal_transactions = None
    try:
        if args.all_reports and not args.download_only:
            from populate_a27 import adjustment_workbook, require_closed_a27

            adjustment = adjustment_workbook()
            require_closed_a27(adjustment)
        for path in run_native_reports(args.date, statements=args.statements, all_reports=args.all_reports,
                                       top_suppliers=args.top_suppliers, top_customers=args.top_customers,
                                       account_transactions=args.account_transactions, vat_returns=args.vat_returns,
                                       aged_summaries=args.aged_summaries,
                                       link_reports=not (args.download_only or args.all_reports),
                                       aged_totals=aged_totals):
            print(f'Saved original Xero file: {path.name}', flush=True)
            if re.fullmatch(r'N9(?: v[0-9]+)?\.xlsx', path.name):
                trial_balance = path
            elif re.fullmatch(r'N8(?: v[0-9]+)?\.xlsx', path.name):
                nominal_transactions = path
    except WorkbookOpenError as error:
        parser.exit(3, f'{error}\n')
    except (OSError, ValueError, PlaywrightError) as error:
        parser.exit(1, f'Xero export failed: {error}\n')
    if args.company_number:
        try:
            download_company(args.company_number, link_reports=False)
        except WorkbookOpenError as error:
            parser.exit(3, f'{error}\n')
        except (OSError, ValueError, PlaywrightError) as error:
            parser.exit(4, f'Companies House export failed: {error}\n')
    if args.all_reports and not args.download_only:
        try:
            from analyse_nominals import analyse_nominals
            from fill_aged_totals import fill_aged_totals
            from fill_working_papers_header import fill_header
            from populate_a27 import populate_a27

            if trial_balance is None:
                raise ValueError('No Trial Balance Excel file was exported for A27.')
            if nominal_transactions is None:
                raise ValueError('No N8 Account Transactions Excel file was exported for N6.')
            populate_a27(trial_balance, adjustment)
            workbook = latest_working_papers(CLIENT_FOLDER)
            fill_header(workbook, DEMO_NAME, args.date,
                        client_code=args.client_code, preparer_initials=args.preparer_initials)
            analyse_nominals(workbook, nominal_transactions, trial_balance, args.date)
            fill_aged_totals(workbook, aged_totals, args.date)
            link_existing_reports()
        except WorkbookOpenError as error:
            parser.exit(3, f'{error}\n')
        except (OSError, ValueError, ArithmeticError) as error:
            parser.exit(1, f'Could not complete A27 or working-paper updates: {error}\n')
    if args.aged_summaries and not args.download_only:
        from fill_aged_totals import fill_aged_totals

        try:
            fill_aged_totals(latest_working_papers(CLIENT_FOLDER), aged_totals, args.date)
        except WorkbookOpenError as error:
            parser.exit(3, f'{error}\n')
        except (OSError, ValueError, ArithmeticError) as error:
            parser.exit(1, f'Could not update aged totals: {error}\n')



if __name__ == '__main__':
    main()
