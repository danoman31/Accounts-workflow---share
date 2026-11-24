# Accounting report workflow

Python scripts for a macOS working-papers workflow. They export original Xero reports, optionally retrieve public Companies House filings, and update a compatible Excel workbook.

This repository contains source code only. The working-papers workbook, firm templates, sample client data, generated reports, Chrome session and credentials are supplied locally. The scripts expect a particular workbook layout and filing structure.

## Set up

Install Python and Google Chrome, then run from the repository root:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install playwright==1.65.2
.venv/bin/python scripts/download_xero_reports.py --sign-in
```

Complete Xero sign-in in the dedicated Chrome window. The current browser workflow selects **Demo Company (UK)**.

## Use the workflow

Put your authorised, compatible templates and client workbook in the expected local directories. The client directory defaults to `outputs/ABC Ltd/`, with a separate subfolder for each working-paper section. The optional API-generated export also expects a local `trial_balance_template.xlsx` at the repository root. Close the working papers and A27 workbook in Excel before running the full workflow.

```bash
.venv/bin/python scripts/download_xero_reports.py --all-reports --date 2026-03-31
```

Use `--help` to see how to include a Companies House number, client code and preparer initials. The Telegram listener is optional and obtains its bot token locally at startup. Keep credentials in `config/` and reports in `outputs/`.
