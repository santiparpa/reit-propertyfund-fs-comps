# Project notes for Claude Code

Project-specific guidance for future Claude sessions working in this folder.
Read this before making changes.

## What this is

A small ETL + dashboard pipeline for Thai REIT / PFPO quarterly financial
statements. See `README.md` for the user-facing overview. The user is the
analyst maintaining it — they refresh it quarterly when new earnings drop.

## The pipeline (single source of truth)

```
raw/*.xlsx  →  scripts/build_db.py  →  data/database.csv
              scripts/build_dashboard.py  →  index.html (data embedded inline)
```

`scripts/build.py` runs both in sequence. **Always re-run it after touching
either build script** — never edit `data/database.csv` or `index.html` directly,
they're regenerated wholesale on every build.

## Hard rules

1. **`index.html` is generated.** The template lives inside `HTML_TEMPLATE` in
   `scripts/build_dashboard.py`. Direct edits to `index.html` are lost on the
   next build. If a redesign is requested, edit the template.
2. **CSV schema is the contract.** `build_db.py` writes 8 columns:
   `symbol, period, period_end, fiscal_period, item, item_path, level, value`.
   `build_dashboard.py` reads exactly that. Don't change one without the other.
3. **`period` is always a calendar quarter** (Jan–Mar → Q1, etc.), derived from
   `period_end`'s month. `fiscal_period` is the company's own Q-label, kept
   alongside for transparency. Seven REITs use non-calendar fiscal years
   (Apr/Jul/Oct starts) — the dashboard surfaces both labels.
4. **Don't break JS DOM hooks.** The dashboard JS depends on specific element
   IDs and class patterns (`#ts-symbols`, `#ts-item`, `#ts-mode`, `#ts-range`,
   `#peer-sort`, `#peer-compare`, `#heat-mode`, `#heat-sort`, `.tab.active`,
   `.pos`/`.neg`, etc.). If you rewrite the HTML template, preserve all of
   these. Full list lives in `build_dashboard.py` near the template — search
   "required IDs" if you're refactoring.
5. **De-duplication is intentional, not a bug.** `Increase (Decrease) In Net
   Assets From Operations` deliberately appears in both Profit & Results and
   Other Comprehensive Income — SETSmart prints it twice and for some stocks
   (AIMIRT, DREIT) the values diverge after restatements. The dedup logic
   keys on `(category, name)` to preserve both. Don't "fix" this by removing
   the duplicate.

## Categorisation logic (lives in `build_dashboard.py`)

Items are routed into four categories independent of the raw item_path indent:

- `Revenue` — path starts with `Revenue >`
- `Other Comprehensive Income` — path starts with `Other Comprehensive Income >`
  (the OCI re-listed Increase/Decrease line) OR `Items That Will [Not] Be
  Subsequently Reclassified...` (OCI subsection items)
- `Profit & Results` — leaf name in `PROFIT_RESULT_NAMES`, OR any sub-item of
  `Other Gains (Losses)`
- `Expenses` — everything else (the actual operating expenses)

Rule ordering matters — rules short-circuit top-to-bottom. Adding a new
category, or a new line item that needs special routing, means editing
`categorize()` + the manual sort lists in `ITEM_ORDER`.

## Adding a new quarter

The pipeline handles new quarters with zero code changes:

1. User drops `{SYMBOL}_FS_{NEW_PERIOD}.xlsx` files into `raw/`.
2. Run `python3 scripts/build.py`.

The CSV picks up the new column from each xlsx's quarter header; the calendar
period is derived from the date range; new periods automatically appear in
every dropdown. Coverage counts auto-update. If a brand-new line item shows
up that isn't in `ITEM_ORDER`, it falls to the end of its category — that's
acceptable; nudge it into the right spot only if the user complains.

## What NOT to add

The user prefers a minimal, self-contained pipeline. Things they have explicitly
rejected or that should not be reintroduced without asking:

- **External servers / databases.** The dashboard is a single HTML file with
  data embedded inline — no fetch(), no CORS, no server required. The SQLite
  prototype was archived for this reason.
- **Heavy build steps.** No bundlers, no transpilers. Plain Python stdlib +
  `openpyxl`. Plain JS + Chart.js from CDN. Inter from Google Fonts.
- **Speculative line items / categories.** If SETSmart didn't emit it, don't
  invent it. The dashboard reflects what's in the source data, not derived
  metrics (yet — they're a possible future addition but ask first).

## Archived work

`archive/old-sqlite-pipeline/` contains a parallel SQLite-based pipeline that
predated the current one. It has features the current pipeline doesn't
(entity metadata with `latestPrice`, source-file provenance, SQL views).
**Don't run it** — it would write to the same `data/` folder and confuse
things. Port features into the current pipeline if/when the user asks.

## Common pitfalls

- The xlsx header row is row ~19 (not fixed) — locate by the first cell that
  matches the quarter-label regex, don't hardcode the row index.
- The `' - '` sentinel in xlsx cells means missing; `to_number()` returns
  `None` for it. Missing values are omitted from the CSV entirely.
- `period_end` is parsed from a `dd/mm/yy` string in the column header. SETSmart
  uses two-digit years → assume `20yy`.
- `select[multiple]` in HTML doesn't honour `flex: 1` reliably unless the
  parent uses `display: contents` on the label (current dashboard does this).
- `Chart.js` instances must be `.destroy()`-ed before recreating them on
  re-render or the canvas accumulates ghost data.

## Useful one-liners

```bash
# Full rebuild
python3 scripts/build.py

# Spot-check a value
python3 -c "import csv; rows=[r for r in csv.DictReader(open('data/database.csv')) if r['symbol']=='CPNREIT' and r['item']=='Total Revenue' and r['period']=='2026Q1']; print(rows)"

# Syntax-check the embedded JS without opening a browser
python3 -c "import re; h=open('index.html').read(); s=[x for x in re.findall(r'<script(?:\s[^>]*)?>(.*?)</script>',h,re.S) if 'tsRender' in x][0]; open('/tmp/c.js','w').write(s)" && node --check /tmp/c.js
```
