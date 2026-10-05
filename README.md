# Financial Provisions and Reversals Manager

A desktop application for tracking financial provisions, reversals, and client records with Excel-backed storage and reporting.

> Public source edition. Configure credentials locally and use empty or synthetic inputs. Company and institution names identify the original integration context; this repository does not claim affiliation or endorsement.

## Capabilities

- Create and manage clients, provisions, and reversals.
- Track provision balances and write-offs.
- Import, update, and export Excel workbooks.
- Provide formatted operational summaries.

## Stack

Python, Flet, Pandas, OpenPyXL, JSON, and Excel workbooks.

## Run locally

Install dependencies from `requirements.txt`, configure a sanitized local workbook, then run `python main.py`. Do not use the committed workbook with real financial operations.

## Empty demo inputs

Run `python scripts/init_demo.py` from the repository root to create an empty local workbook and tax configuration. Existing files are preserved. Add demonstration records through the application; real financial inputs stay local.
