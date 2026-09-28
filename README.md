# Sales Data Cleaning & Exploratory Analysis

Cleaning a messy e-commerce sales export (486 orders, 3 months, fictional
Egyptian online clothing store) and turning it into a clean dataset, an
exploratory analysis, and a client-ready Excel report.

> **Note on the data:** this is a **simulated dataset**, built to mimic the
> kind of inconsistencies a real operations export actually has — not a
> pre-cleaned Kaggle dataset, and not figures from a real business. The
> cleaning problems (mixed formats, duplicates, missing values) are
> realistic; the business numbers are illustrations of the analysis.

## Problem

The raw file had the kind of inconsistencies a real operations sheet
actually has:
- Order dates in 5+ formats, including a **genuinely ambiguous** one
  (`04.07.2026` could be 4 July or 4 April — see below)
- City names in mixed Arabic/English spellings, casing, and a `'-'`
  placeholder for missing values
- Prices as numbers, currency-labeled strings, and comma-formatted strings
  (`'EGP 950'`, `'950 جنيه'`, `'1,100'`)
- Phone numbers in local, `+20`, and malformed (dropped leading zero) formats
- Product/payment/status values with inconsistent casing and spelling
- 36 fully duplicated rows

## The ambiguous-date problem

Some dates are written as `D.M.Y` and others as `M.D.Y`, with no way to tell
which from the string alone when both numbers are ≤ 12 (e.g. `04.07.2026`).
`clean_date` resolves this using the file's own data: it parses every
unambiguous date first, uses those to find the file's true date range, then
picks whichever reading of an ambiguous date falls inside that range.
Details and reasoning: [`src/cleaning.py`](src/cleaning.py) and
[`notebooks/analysis.ipynb`](notebooks/analysis.ipynb).

## Pipeline

```
Raw Data → Cleaning → EDA → Visualization → Excel report
```

## Key Insights

- **Tops** is the leading product category at **31.1%** of all orders.
- By revenue, the top 5 products are **Jacket, Abaya, Sneakers, Dress,
  Jeans** — driven mainly by price, since unit sales are nearly flat
  (69–76 units) across products.
- **Cairo** leads on both orders (118) and revenue (~95K EGP), followed by
  Giza and Alexandria; order volume roughly tracks city population.
- **Cash on Delivery** remains the top payment method (31.6%), but roughly
  two-thirds of orders use a digital/pre-paid method (Instapay + Credit
  Card + Vodafone Cash combined).
- **~30% of orders are cancelled or returned** (15.3% / 14.9%) — a strong
  signal to investigate further before treating gross revenue as realized
  revenue.

Full analysis, charts, and reasoning: [`notebooks/analysis.ipynb`](notebooks/analysis.ipynb)
Client-facing report: [`reports/cleaned_sales_report.xlsx`](reports/cleaned_sales_report.xlsx)

## Repo Structure

```
├── data/
│   ├── raw/                    # original messy Excel export
│   └── cleaned/                 # cleaned CSV output
├── src/
│   ├── cleaning.py              # one documented function per column
│   └── report.py                # builds the 3-sheet Excel report (live formulas)
├── notebooks/
│   └── analysis.ipynb           # audit → cleaning → EDA → charts → insights
├── reports/
│   └── cleaned_sales_report.xlsx  # Report / Cleaned Data / Cleaning Notes
├── images/                      # exported chart PNGs
└── requirements.txt
```

## Missing-Value Policy

| Column | Decision | Reason |
|---|---|---|
| `City`, `Qty`, `Price` | Left blank | Feed numeric aggregations; a guessed value would silently corrupt sums/averages |
| `Phone`, `Customer Name` | Filled with a text placeholder | Never used in a calculation — more readable than a blank in a report |

Full counts and per-column notes: `Cleaning Notes` sheet in the Excel report.

## How to Run

```bash
pip install -r requirements.txt
jupyter notebook notebooks/analysis.ipynb
```

## Tech Stack

Python · pandas · NumPy · Matplotlib · openpyxl · Jupyter

## Limitations & Next Steps

- Missing values were not imputed/guessed — totals are slightly understated
  rather than wrong (about 5.6% of orders have no price).
- Next: relational analysis with SQL (Project 2), then a model to predict
  order cancellation/return risk (Project 3).
