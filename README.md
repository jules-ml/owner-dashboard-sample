# Sample: messy business data → owner dashboard

**SAMPLE: synthetic data for a fictional business.** The company name is Lantern Key Pool & Air, a made-up pool and HVAC shop in the Delray Beach / Palm Beach County, FL area. There are no real customers, invoices, or results. Emails use the `.invalid` TLD. Phone numbers use the 555 prefix reserved for fiction.

This is a public portfolio sample for Giulio Piccinonna’s “messy business data → clean owner dashboard” pitch.

As-of date for accounts-receivable aging: **2026-11-25** (America/New_York). Source activity runs from 2026-01-06 through 2026-11-20.

## What is in this folder

```
generate_data.py      Builds the deliberately messy source exports
clean_data.py         Pandas cleaning + validation pipeline
app.py                Streamlit owner dashboard
requirements.txt
data/raw/             Messy exports (the “before”)
  invoices_payments.csv
  jobs_expenses.xlsx   (Jobs Export + Expenses; headers are not on row 1)
  crm_leads.csv
  README.txt
data/clean/           Clean tables, KPI extract, and the data-quality log
```

## How to run

```bash
python3 -m pip install -r requirements.txt
python3 generate_data.py    # optional; raw files are already checked in
python3 clean_data.py
streamlit run app.py
```

The dashboard loads tables from `data/clean/` (written by `clean_data.py`). If those files are missing, `app.py` runs the cleaning step on startup. It shows:

- Revenue and gross profit by month
- Profit by service line
- AR aging and a “chase first” list
- Cash in vs cash out
- A data-quality panel with before/after counts and flagged rows

Revenue is pre-tax invoice totals. Gross profit is revenue minus direct job labor and materials. Cash out treats those job costs as paid in the job month, plus overhead (rent, insurance, fuel, truck lease, and similar). That is a documented simplification.

## What the cleaning step does

The pipeline does not use a hardcoded customer answer key. It:

- Skips preamble rows and finds the real header in each export
- Parses mixed date formats (US, day-first when unambiguous, month names, Excel serials) and mixed currency formats (US and European)
- Standardizes service-line aliases
- Merges duplicate customers from name variants, nicknames, punctuation, and shared email or phone
- Collapses payment-split rows into one invoice plus many payments
- Drops exact duplicate rows
- Flags unparseable amounts and dates, overpayments, typo dates, and jobs with no matching customer, instead of guessing

Re-run `python3 clean_data.py` after changing the raw files. The log is written to `data/clean/data_quality_log.md`.


## Deploy on Streamlit Community Cloud

This repo is set up to run `app.py` from the repository root on the `main` branch, with Python 3.12 (see `.python-version`). After the public repo exists, connect it in Streamlit Community Cloud and set the main file path to `app.py`. The numbers in the app are synthetic sample data, not real business results.
