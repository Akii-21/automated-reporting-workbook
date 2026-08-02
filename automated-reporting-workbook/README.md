# Automated Reporting Workbook

A config-driven Python pipeline that turns a raw sales/orders CSV into a **live Excel workbook** and a **shareable HTML dashboard** — automatically, on a schedule, via GitHub Actions.

Built as a data analytics internship project to demonstrate: data validation, KPI logic, spreadsheet formula design, and basic CI/CD automation.

## Why this exists

Most "reporting" tasks in an analyst role are the same five steps repeated every week: pull the data, clean it, recompute the same KPIs, refresh the same charts, and send the same workbook to someone. This project automates that loop end-to-end so it takes **one command** (or zero — it also runs itself on a schedule).

## What it produces

Running the pipeline on a sales dataset generates two artifacts in `reports/`:

| File | What it is |
|---|---|
| `sales_report.xlsx` | A multi-sheet workbook — **Dashboard**, **Raw Data**, **Category Summary**, **Regional Summary**, **Monthly Trend** — where every summary number is a live `SUMIF`/`SUMPRODUCT` formula referencing the Raw Data sheet, not a pasted-in value. Edit a row in Raw Data, hit recalculate, and the whole report (including the native Excel charts) updates. |
| `sales_report.html` | A single, self-contained HTML dashboard (charts embedded as images) for quick sharing or "print to PDF" — no Excel required. |

A sample of both, generated from the included demo dataset, is already committed in [`/reports`](reports) so you can preview the output without running anything.

## How it works

```
data/your_export.csv ──▶ data_loader.py ──▶ kpi_engine.py ──┬──▶ excel_report.py ──▶ sales_report.xlsx
                         (validate, clean)   (aggregate)     └──▶ html_report.py  ──▶ sales_report.html
```

The pipeline doesn't assume a fixed set of column names. `config.yaml` maps *your* CSV's headers onto the fields the pipeline needs (`date`, `sales`, `category`, `region`, ...), so pointing it at a different dataset — from Kaggle, an export from your company's CRM, whatever — is a config edit, not a code change.

## Project structure

```
automated-reporting-workbook/
├── generate_report.py         # CLI entry point
├── config.yaml                 # column mapping + report settings
├── requirements.txt
├── data/
│   └── sample_sales_data.csv   # demo dataset (synthetic, Kaggle-style retail sales)
├── src/
│   ├── data_loader.py          # load + validate CSV against config
│   ├── kpi_engine.py           # all aggregation logic (single source of truth for "what to calculate")
│   ├── excel_report.py         # builds the .xlsx (formulas + native charts)
│   └── html_report.py          # builds the .html (matplotlib charts, embedded)
├── reports/                     # generated output (sample committed)
└── .github/workflows/
    └── generate_report.yml     # re-runs the pipeline on push + weekly schedule
```

## Running it

```bash
pip install -r requirements.txt
python generate_report.py --input data/sample_sales_data.csv
```

Output lands in `reports/sales_report.xlsx` and `reports/sales_report.html`.

To point it at your own data:
1. Drop your CSV in `data/`.
2. Edit the `columns:` block in `config.yaml` so the right-hand side matches your CSV's actual headers.
3. Run `python generate_report.py --input data/your_file.csv`.

Useful flags: `--skip-excel`, `--skip-html`, `--outdir <path>`, `--config <path>`.

## Automation

[`.github/workflows/generate_report.yml`](.github/workflows/generate_report.yml) reruns the pipeline and commits the refreshed report whenever `data/` or the pipeline code changes, and every Monday morning on a cron schedule — simulating a recurring "Monday morning sales report" that a real analyst would otherwise generate by hand.

## Design notes

- **Formulas over hardcoded values.** Every number in the Excel summary sheets is computed by a spreadsheet formula (`SUMIF`, `SUMPRODUCT`, `IFERROR`) that references the Raw Data sheet — not a Python-computed number pasted into a cell. This is the difference between a "report" and a workbook someone could actually keep using.
- **Config over hardcoding.** Column names, report title, and top-N cutoffs live in `config.yaml`, not scattered through the code, so the same pipeline works on a different dataset without touching `src/`.
- **`kpi_engine.py` is the single source of truth** for what gets calculated; both the Excel and HTML builders call into it, so the two outputs can never disagree with each other.

## Sample dataset

`data/sample_sales_data.csv` is a synthetic 3,000-row retail sales dataset (Order ID, Date, Region, Category, Sub-Category, Segment, Quantity, Sales, Profit) generated to resemble common Kaggle retail-sales datasets, with mild seasonality built in around Nov/Dec. Swap it for a real Kaggle dataset (e.g. Superstore Sales, or any orders-level CSV) by updating `config.yaml`.

## Stack

Python · pandas · openpyxl · matplotlib · PyYAML · GitHub Actions
