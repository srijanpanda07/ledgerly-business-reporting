# Ledgerly — Business Report Automator

[Live demo](https://ledgerly-sales-reporting.streamlit.app/) · [Source code](https://github.com/srijanpanda07/ledgerly-business-reporting)

Ledgerly turns a sales CSV or Excel workbook into an interactive overview, data-quality review, and shareable Excel or PDF report. The interface uses a quiet sage, charcoal, and warm-neutral palette designed for readable business reporting.

## What it does

- Reads CSV, XLSX, and XLSM files, with a fictional sample dataset for trying the app.
- Maps sales, date, category, product, order, quantity, and unit-price columns.
- Calculates revenue from a sales column or from quantity × unit price.
- Supports date and category filters shared across the workspace.
- Shows sales totals, monthly trends, category and product summaries, and data-quality checks.
- Exports filtered rows as CSV and reports as Excel workbooks or PDF summaries.

## Run locally

In PowerShell from the project folder:

    py -m venv .venv
    .\.venv\Scripts\python.exe -m pip install -r requirements.txt
    .\.venv\Scripts\python.exe -m streamlit run app.py

The app opens at http://localhost:8501.

## Deploy to Streamlit Community Cloud

The public demo is running at [ledgerly-sales-reporting.streamlit.app](https://ledgerly-sales-reporting.streamlit.app/). The source repository is connected to Streamlit Community Cloud, which redeploys changes pushed to the main branch.

The app does not require API keys or secrets. A public deployment lets visitors upload files to the running app, so use fictional or anonymised data. Uploaded files are processed by the app; this project has no login, database, or private report storage.

## Data notes

The sample data is fictional. This app provides descriptive sales reporting, not accounting advice. Sales are not profit; refunds, taxes, discounts, and cost data may need separate treatment. Date parsing depends on the source format, and duplicate detection compares exact rows only. Review the data-quality page and verify exported figures against your source system.
