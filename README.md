# Financial Provision Manager

A desktop application for tracking financial provisions, reversals, and client records with Excel-backed storage and reporting.


## Capabilities

- Create and manage clients, provisions, and reversals.
- Track provision balances and write-offs.
- Import, update, and export Excel workbooks.
- Provide formatted operational summaries.

## Stack

Python, Flet, Pandas, OpenPyXL, JSON, and Excel workbooks.

## Run locally

Run `pip install -r requirements.txt`, initialize the empty inputs with `python scripts/init_demo.py`, then run `python main.py`. The public template contains column headers only.

## Empty demo inputs

Run `python scripts/init_demo.py` from the repository root to create an empty local workbook and tax configuration. Existing files are preserved. Add demonstration records through the application; real financial inputs stay local.


## Snapshot review

Setup and capabilities are based on the included code. External services and operating-system-specific integrations were not exercised during this publication review. Use synthetic data in a separate local environment.

Git-triggered Vercel deployment is disabled in the included web configurations. This update publishes source and documentation only.

## License

No new distribution license is granted for this corporate source edition. Public visibility does not grant reuse rights; obtain permission from the relevant rights holder. Existing third-party notices remain applicable.

## Verification record

See [snapshot checks and remaining limitations](docs/VERIFICATION.md).
