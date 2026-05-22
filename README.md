# Thai REIT / PFPO Financial Statements

A small, self-contained pipeline that turns SETSmart-exported quarterly
financial-statement workbooks into a normalised database and a single-file
HTML dashboard for time-series, peer-comparison, and ad-hoc analysis across
56 Thai REITs and property funds.

## Folder layout

```
.
├── raw/                          ← drop SETSmart .xlsx exports here
│   └── {SYMBOL}_FS_{PERIOD}.xlsx     (56 files, one per REIT/PFPO)
├── data/
│   └── database.csv              ← normalised long-format DB (generated)
├── scripts/
│   ├── build_db.py               ← xlsx → CSV
│   ├── build_dashboard.py        ← CSV → index.html (data embedded)
│   └── build.py                  ← runs both, single command
├── index.html                    ← the dashboard (open in any browser)
├── WORKFLOW.md                   ← quarterly download + rename procedure
├── CLAUDE.md                     ← project notes for Claude Code sessions
├── archive/
│   └── old-sqlite-pipeline/      ← superseded SQLite-based prototype (kept for reference)
└── README.md                     ← this file
```

## Quarterly refresh

When new earnings drop:

1. Download the new SETSmart exports and place them into `raw/`. Use
   `WORKFLOW.md` for the rename step that turns `FinancialStatement (n).xlsx`
   into `{SYMBOL}_FS_{PERIOD}.xlsx`.
2. Rebuild everything in one command:

```bash
python3 scripts/build.py
```

3. Open `index.html` — the new quarter is automatically pulled into every
   view. Re-running the build never modifies the raw xlsx files.

## What the dashboard shows

Four views (left sidebar):

- **Time series** — quarterly trend for multiple REITs side-by-side. Modes:
  Absolute, QoQ %, YoY %, Index = 100 @ first quarter. Range filter for
  1y / 2y / 3y / all. Latest-snapshot table below shows latest value, QoQ,
  YoY, 4-year average per selected stock.
- **Peer comparison** — ranked horizontal bar chart across all 56 stocks for
  a chosen line item + quarter. Compare modes: Absolute, vs. prior Q %,
  YoY %, or % of Total Revenue.
- **QoQ / YoY heatmap** — 56 stocks × 16 calendar quarters colour-scaled
  matrix. Spot outliers and trends fast.
- **Raw data** — every line item in long format, filterable by symbol,
  category, item name, period, and minimum magnitude. Sortable columns.

## Data model (CSV)

`data/database.csv` is the canonical normalised database. Long format, eight
columns:

| Column          | Description |
|-----------------|---|
| `symbol`        | Stock ticker (e.g. CPNREIT) |
| `period`        | **Calendar** quarter (e.g. `2026Q1`) — derived from `period_end` month |
| `period_end`    | ISO date the quarter ended (e.g. `2026-03-31`) |
| `fiscal_period` | The company's own Q-label (differs from `period` for Apr/Jul/Oct-start fiscal years) |
| `item`          | Leaf line-item name (e.g. `Total Revenue`) |
| `item_path`     | Full hierarchical path (e.g. `Revenue > Total Revenue`) |
| `level`         | Indent depth in the SETSmart source (2 = main item, 3 = sub-item) |
| `value`         | Numeric value in k.Baht |

Currently 13,973 rows from 56 stocks × ~30 line items × up to 16 calendar
quarters (2022 Q2 → 2026 Q1).

## Non-calendar fiscal years

Seven stocks use non-calendar fiscal years. The pipeline auto-detects this
from each xlsx's column-header date range and realigns everything to
**calendar quarters**, keeping the company's own fiscal label in
`fiscal_period`:

| Fiscal year starts | Symbols |
|---|---|
| April | IMPACT, TIF1, WHABT |
| July  | LUXF, SSPF |
| October | FTREIT, GVREIT |

The dashboard shows a banner identifying these symbols and surfaces the fiscal
label in tooltips, the latest-snapshot table, and the raw-data viewer.

## Line-item categorisation

Items are grouped into four logical categories — independent of the SETSmart
indent which is sometimes misleading (Net Investment Income, Other Gains, etc.
sit under "Expenses" by indent but are subtotals, not expenses):

- **Revenue** — Total Revenue and its components
- **Expenses** — actual operating expenses (Management Fee, Finance Costs, ...)
- **Profit & Results** — Net Investment Income, Other Gains (Losses) and sub-items,
  Increase (Decrease) In Net Assets From Operations, Total Comprehensive Income
- **Other Comprehensive Income** — items in the OCI section of the statement
  (Currency Translation Adjustments, Remeasurement Of Employee Benefits, etc.)

`Increase (Decrease) In Net Assets From Operations` deliberately appears in
both **Profit & Results** and **Other Comprehensive Income** — SETSmart prints
the value twice and for some REITs (AIMIRT, DREIT) the two values diverge
after restatements. The dropdown disambiguates with `· P&L` / `· OCI` suffixes.

## Dependencies

- Python 3.10+
- `openpyxl` — `pip install openpyxl`
- Any modern browser (the dashboard uses Chart.js via CDN and Inter from
  Google Fonts; works offline too with cached assets)

## Universe (56 symbols)

```
AIMCG, AIMIRT, ALLY, AMATAR, AXTRART, BAREIT, BOFFICE, B-WORK,
CPNCG, CPNREIT, CPTREIT, CTARAF, DREIT, FTREIT, FUTURERT,
GAHREIT, GROREIT, GVREIT, HPF, HYDROGEN, IMPACT, INETREIT,
ISSARA, KPNREIT, KTBSTMR, LHHOTEL, LHRREIT, LHSC, LUXF, MII,
MIPF, MJLF, MNIT, MNIT2, MNRF, M-PAT, M-STOR, POPF, PROSPECT,
QHBREIT, QHHRREIT, QHOP, SIRIPRT, SPRIME, SRIPANWA, SSPF, SSTRT,
TIF1, TLHPF, TNPF, TPRIME, TTLPF, TU-PF, WHABT, WHAIR, WHART
```

Update `scripts/build_db.py` (or just drop new files into `raw/`) when REITs
are added or delisted from the SET.
