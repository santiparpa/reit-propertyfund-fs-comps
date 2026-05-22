# Archived: old SQLite-based pipeline

These files implement an earlier dashboard pipeline that has been superseded
by the canonical CSV + self-contained HTML pipeline in the project root.

They are kept here for reference only — nothing in the current workflow reads
or writes any of them.

## What's here

| File | What it was |
|---|---|
| `scripts/build_financial_database.py` | Built a normalised SQLite database and a browser-ready JS data export from the raw xlsx workbooks. |
| `data/financials.sqlite` | SQLite database: `entities`, `periods`, `financial_items`, `financial_facts`, `source_files` tables plus `v_financial_facts` / `v_period_changes` views. |
| `data/financials_dashboard_data.js` | Pre-computed JS payload (`window.FINANCIAL_DATA = {...}`) consumed by the old HTML dashboard. |
| `reit_financial_dashboard.html` | Standalone dashboard that loaded the JS payload above and rendered time-series + peer comparison charts. |

## Useful features it had (that the current pipeline doesn't)

- **Entity metadata**: captured `latestPrice`, `market`, `sector`, `industry`
  from cells B9–B12 of each workbook (the current pipeline only captures the
  income-statement values).
- **Source-file provenance**: tracked filename, sha and import timestamp per
  workbook in `source_files`.
- **Stable item keys**: each line item got a canonical `item_key` (e.g.
  `total_revenue`) instead of relying on the raw item name string.
- **SQL views**: `v_financial_facts` joined the four core tables for ad-hoc
  queries, and `v_period_changes` pre-computed QoQ/YoY deltas.

If any of these become valuable again, the source is right here — port what you
need into the current pipeline (`scripts/build_db.py` + `scripts/build_dashboard.py`).

## Why it was archived

The user adopted the simpler CSV + single-file HTML pipeline (`scripts/build_db.py`
→ `data/database.csv` → `scripts/build_dashboard.py` → `index.html`) as their
daily-driver dashboard. Two parallel pipelines covering the same data was
confusing, so this one was archived rather than deleted in case the SQLite shape
or extra metadata is useful later.
