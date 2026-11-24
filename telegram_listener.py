"""Accept paired Telegram commands for the UHY Demo Company export."""
import argparse
from datetime import date, datetime
from getpass import getpass
import json
import re
import secrets
from subprocess import TimeoutExpired, run
import sys
import time
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from download_xero_reports import xero_tls_context
from local_credentials import read_credential, write_credential
from project_paths import PROJECT_ROOT, SCRIPTS_FOLDER


USAGE = ('Send /reports demo-uk YYYY-MM-DD or DD/MM/YYYY using the financial year-end '
         'for Demo Company (UK). Optionally add the Companies House number, client=ABC01 and preparer=DS.')


class TelegramError(Exception):
    def __init__(self, code, method):
        self.code = code
        super().__init__(f'Telegram returned error {code} for {method}.')


def telegram(token, method, **fields):
    url = f'https://api.telegram.org/bot{token}/{method}'
    try:
        with urlopen(Request(url, data=urlencode(fields).encode()),
                     context=xero_tls_context(), timeout=35) as response:
            payload = json.load(response)
    except HTTPError as error:
        raise TelegramError(error.code, method) from None
    except (OSError, ValueError):
        raise TelegramError('network or response failure', method) from None
    if not payload['ok']:
        raise TelegramError(payload['error_code'], method)
    return payload['result']


def export_command(message):
    match = re.fullmatch(
        r'/(reports|tb)\s+demo-uk\s+([0-9]{4}-[0-9]{2}-[0-9]{2}|[0-9]{2}/[0-9]{2}/[0-9]{4})'
        r'(?:\s+(?P<number>[A-Za-z0-9]{8}))?'
        r'(?:\s+client=(?P<client>[A-Za-z0-9_.-]+))?'
        r'(?:\s+preparer=(?P<preparer>[A-Za-z0-9_.-]+))?',
        message,
    )
    if not match or (match[1] == 'tb' and any(match.group(name) for name in ('number', 'client', 'preparer'))):
        return USAGE
    try:
        raw_date = match[2]
        as_at = (datetime.strptime(raw_date, '%d/%m/%Y').date()
                 if '/' in raw_date else date.fromisoformat(raw_date))
    except ValueError:
        return USAGE
    command = [
        str(PROJECT_ROOT / '.venv' / 'bin' / 'python'), '-B',
        str(SCRIPTS_FOLDER / 'download_xero_reports.py'), '--date', as_at.isoformat(),
    ]
    if match[1] == 'reports':
        command.append('--all-reports')
    for name, flag in [('number', '--company-number'), ('client', '--client-code'),
                       ('preparer', '--preparer-initials')]:
        if value := match.group(name):
            command.append(f'{flag}={value}')
    try:
        result = run(command, capture_output=True, text=True, timeout=300)
    except (OSError, TimeoutExpired) as error:
        print(f'Xero download failed: {error}', file=sys.stderr, flush=True)
        return 'Xero download failed. Check Terminal for details.'
    summary = result.stdout.strip()
    if result.returncode in (3, 4):
        return '\n'.join(filter(None, (summary, result.stderr.strip())))
    if result.returncode:
        print(result.stderr.strip(), file=sys.stderr, flush=True)
        failure = ('Xero download did not finish. Check Terminal. If Xero asks you to sign in, '
                   'run .venv/bin/python scripts/download_xero_reports.py --sign-in.')
        return '\n'.join(filter(None, (summary, failure)))
    return summary or 'Xero returned no downloads. Check Terminal.'



def listen(token, *, persist_token=False):
    username = telegram(token, 'getMe')['username']
    if persist_token:
        write_credential('telegram.json', {'bot_token': token})
    pairing_code = '/pair ' + secrets.token_urlsafe(12)
    print(f'Open https://t.me/{username} in Telegram and send this command privately:',
          pairing_code, sep='\n', flush=True)
    owner = None
    offset = 0
    while True:
        try:
            updates = telegram(token, 'getUpdates', offset=offset, timeout=20,
                               allowed_updates='["message"]')
        except TelegramError as error:
            if error.code in (401, 404, 409):
                raise
            print(error, file=sys.stderr, flush=True)
            time.sleep(5)
            continue
        for update in updates:
            offset = update['update_id'] + 1
            message = update.get('message', {})
            chat = message.get('chat', {})
            sender = message.get('from', {}).get('id')
            if chat.get('type') != 'private' or sender is None or chat.get('id') != sender:
                continue
            command = message.get('text', '').strip()
            if owner is None and secrets.compare_digest(command.encode(), pairing_code.encode()):
                owner = sender
                reply = 'Paired. ' + USAGE
            elif owner == sender:
                reply = export_command(command)
            else:
                continue
            try:
                telegram(token, 'sendMessage', chat_id=sender, text=reply)
            except TelegramError as error:
                print(error, file=sys.stderr, flush=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--replace-token', action='store_true', help='Save a new Telegram bot token')
    args = parser.parse_args()
    try:
        saved = None if args.replace_token else read_credential('telegram.json')
        token = saved['bot_token'] if saved is not None else getpass('Paste your Telegram bot token (hidden): ').strip()
        if not re.fullmatch(r'[0-9]+:[A-Za-z0-9_-]+', token):
            sys.exit('That does not look like a Telegram bot token.')
        listen(token, persist_token=saved is None)
    except (TelegramError, OSError, ValueError, KeyError) as error:
        sys.exit(str(error))
    except KeyboardInterrupt:
        print('\nListener stopped.')


if __name__ == '__main__':
    main()
