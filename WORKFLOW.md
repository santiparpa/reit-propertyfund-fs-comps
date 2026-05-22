# Workflow: Standardize Thai REIT/PFPO Financial Statement Downloads

Quarterly task — rename SETSmart-exported `FinancialStatement (n).xlsx` files
using the stock symbol in cell **B3**, then reconcile against the master
list of REIT/PFPO tickers to find which exports failed to download.

## Inputs

1. **Folder of downloads** — typically your browser Downloads folder
   containing files named `FinancialStatement.xlsx`, `FinancialStatement (2).xlsx`, ...
2. **Period label** — e.g. `1Q2026`
3. **Master ticker list** — the canonical universe of REITs + PFPOs to check against
   (currently 56 symbols: see `Master list` section below)

## Output

- Files renamed and moved into `raw/` of this project as:
  `{SYMBOL}_FS_{PERIOD}.xlsx` (e.g. `CPNREIT_FS_1Q2026.xlsx`)
- A diff report listing missing symbols (failed downloads to re-export)

## Steps

### 1. Verify the folder contents

```powershell
ls "<folder>"
```

Expect a set of `FinancialStatement*.xlsx` files. Each file's worksheet has
the stock symbol in cell **B3** on the active sheet (`Sheet1`).

### 2. Rename every file using its B3 symbol and move to `raw/`

Run from the folder of downloads. Adjust `RAW` to point at this project's
`raw/` directory.

```python
import openpyxl, glob, os, shutil
from pathlib import Path

PERIOD = '1Q2026'  # change per quarter
RAW = Path('/path/to/reit&pfpo-financialstatement-1q2026/raw')

renames = []
for f in sorted(glob.glob('FinancialStatement*.xlsx')):
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    symbol = str(wb.active['B3'].value).strip()
    wb.close()
    renames.append((f, RAW / f'{symbol}_FS_{PERIOD}.xlsx'))

# Collision check before any move
targets = [t for _, t in renames]
dupes = {t for t in targets if targets.count(t) > 1}
assert not dupes, f'Duplicate target names: {dupes}'

RAW.mkdir(parents=True, exist_ok=True)
for old, new in renames:
    if new.exists():
        print(f'SKIP (target exists): {new}')
        continue
    shutil.move(old, new)
    print(f'{old}  ->  {new}')
```

Notes:
- `read_only=True, data_only=True` reads cached values without loading
  formulas — fast and safe.
- Always do the collision check **before** renaming any file. Two files
  with the same B3 value should halt the run, not silently overwrite.
- Symbols may contain hyphens (e.g. `TU-PF`, `M-PAT`, `M-STOR`, `B-WORK`) —
  do not sanitize them.

### 3. Diff against the master ticker list

Run from the project root.

```python
expected = {
    'AIMCG','AIMIRT','ALLY','AMATAR','AXTRART','BAREIT','BOFFICE','B-WORK',
    'CPNCG','CPNREIT','CPTREIT','CTARAF','DREIT','FTREIT','FUTURERT',
    'GAHREIT','GROREIT','GVREIT','HPF','HYDROGEN','IMPACT','INETREIT',
    'ISSARA','KPNREIT','KTBSTMR','LHHOTEL','LHRREIT','LHSC','LUXF','MII',
    'MIPF','MJLF','MNIT','MNIT2','MNRF','M-PAT','M-STOR','POPF','PROSPECT',
    'QHBREIT','QHHRREIT','QHOP','SIRIPRT','SPRIME','SRIPANWA','SSPF','SSTRT',
    'TIF1','TLHPF','TNPF','TPRIME','TTLPF','TU-PF','WHABT','WHAIR','WHART',
}

import glob, re
PERIOD = '1Q2026'  # change per quarter
have = set()
for f in glob.glob(f'raw/*_FS_{PERIOD}.xlsx'):
    m = re.search(r'([^/\\]+)_FS_', f)
    if m: have.add(m.group(1))

missing = sorted(expected - have)
extra   = sorted(have - expected)
print(f'Expected: {len(expected)} | Have: {len(have)} | Missing: {len(missing)}')
print('Missing :', ', '.join(missing) or 'none')
print('Extra   :', ', '.join(extra)   or 'none')
```

### 4. Re-export missing symbols

Re-download from SETSmart for each symbol in the `Missing` list. Browsers
throttle bulk downloads, so:
- export in batches of ~10 with a few seconds between clicks, **or**
- allow the "site wants to download multiple files" permission prompt up front

Re-run steps 2–3 after each batch until `Missing: none`.

## Building the analysis dashboard

After the xlsx files are renamed (steps 1–4 above), regenerate the standardized
database and dashboard in one command:

```bash
python scripts/build.py
```

This runs two stages, both idempotent and re-runnable on every refresh:

1. `scripts/build_db.py` — parses every `*_FS_*.xlsx` in `raw/` and writes
   `data/database.csv` in long format (`symbol, period, period_end,
   fiscal_period, item, item_path, level, value`). Period is the **calendar**
   quarter derived from `period_end`'s month, so REITs with non-calendar fiscal
   years (IMPACT, TIF1, WHABT, LUXF, SSPF, FTREIT, GVREIT) are auto-realigned.
   The original company-fiscal label is preserved in `fiscal_period`.
   Hierarchy is inferred from the indent in column A (1 / 4 / 6 spaces →
   level 1 / 2 / 3). The `' - '` sentinel and blanks become missing values
   (omitted from the CSV).
2. `scripts/build_dashboard.py` — reads the CSV and emits a single
   self-contained `index.html` (data embedded as JSON, ~360 KB). Open it
   directly — no local server required.

The dashboard has four views: time-series per stock (abs / QoQ / YoY / indexed),
peer comparison for a single quarter (abs / QoQ / YoY / % of revenue), a QoQ/YoY
heatmap matrix, and a filterable raw-data viewer with categorised line items
(Revenue / Expenses / Profit & Results / Other Comprehensive Income).

### Adding a new quarter

1. Drop the renamed `*_FS_<period>.xlsx` files into `raw/` (use steps 1–4 above
   to standardize filenames).
2. Run `python scripts/build.py`.
3. Open `index.html` — newer calendar quarters appear automatically in every
   dropdown.

No schema changes needed even if new line items show up: build_db.py preserves
every item it encounters, and the dashboard groups items by category with
universal items first within each group.

## Master list (56 symbols, as of 1Q2026)

```
AIMCG, AIMIRT, ALLY, AMATAR, AXTRART, BAREIT, BOFFICE, B-WORK,
CPNCG, CPNREIT, CPTREIT, CTARAF, DREIT, FTREIT, FUTURERT,
GAHREIT, GROREIT, GVREIT, HPF, HYDROGEN, IMPACT, INETREIT,
ISSARA, KPNREIT, KTBSTMR, LHHOTEL, LHRREIT, LHSC, LUXF, MII,
MIPF, MJLF, MNIT, MNIT2, MNRF, M-PAT, M-STOR, POPF, PROSPECT,
QHBREIT, QHHRREIT, QHOP, SIRIPRT, SPRIME, SRIPANWA, SSPF, SSTRT,
TIF1, TLHPF, TNPF, TPRIME, TTLPF, TU-PF, WHABT, WHAIR, WHART
```

Update this list when REITs are added/delisted on the SET.

## Skill conversion notes

To turn this into a Claude Code skill (via `/financial-analysis:skill-creator`):

- **Trigger phrases**: "rename REIT financials", "standardize quarterly REIT downloads",
  "diff REIT exports against master list", "process SETSmart financial statements"
- **Required args**: `folder_path`, `period` (e.g. `1Q2026`)
- **Optional args**: `master_list_path` (override the embedded default)
- **Tool needs**: `Bash` (Python with openpyxl), `Write` (to emit a diff report)
- **Safety**:
  - collision check before any rename
  - never auto-delete; only rename and report
  - halt if any B3 cell is empty or non-string
