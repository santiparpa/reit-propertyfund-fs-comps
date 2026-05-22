"""
build_dashboard.py
------------------
Read data/database.csv and emit a single self-contained index.html in the
project root. The HTML embeds the data as JSON and provides four views:

  1. Time Series        — per-stock trend of one or more line items
  2. Peer Comparison    — ranked bar chart across all REITs for one item+quarter
  3. QoQ/YoY Heatmap    — symbol × quarter matrix, colour-scaled
  4. Raw Data Viewer    — filterable, sortable long-format table

Re-runnable: any time you refresh data/database.csv, run this to regenerate
index.html with the new data baked in.
"""

from __future__ import annotations

import csv
import json
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CSV_PATH = ROOT / "data" / "database.csv"
LABELS_PATH = ROOT / "data" / "reit&pfpo-name-list.xlsx"
HTML_PATH = ROOT / "index.html"


def load_csv() -> list[dict]:
    if not CSV_PATH.exists():
        sys.exit(f"Missing {CSV_PATH}. Run scripts/build_db.py first.")
    with CSV_PATH.open(encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def load_labels() -> dict[str, dict]:
    """Read symbol → {type, industry} from the labels xlsx.

    Header row is (Name, Type, Industry). Missing file is non-fatal — the
    dashboard still works, just without the filters populated.
    """
    if not LABELS_PATH.exists():
        print(f"  (no labels file at {LABELS_PATH.relative_to(ROOT)} — skipping)")
        return {}
    try:
        import openpyxl  # type: ignore
    except ImportError:
        sys.exit("openpyxl is required to read labels. Install: pip install openpyxl")
    wb = openpyxl.load_workbook(LABELS_PATH, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {}
    header = [str(c).strip().lower() if c else "" for c in rows[0]]
    try:
        i_name = header.index("name")
        i_type = header.index("type")
        i_ind = header.index("industry")
    except ValueError:
        sys.exit(f"Labels file header must contain Name/Type/Industry; got {header}")
    out: dict[str, dict] = {}
    for r in rows[1:]:
        if not r or r[i_name] is None:
            continue
        sym = str(r[i_name]).strip()
        if not sym:
            continue
        out[sym] = {
            "type": str(r[i_type]).strip() if r[i_type] is not None else "",
            "industry": str(r[i_ind]).strip() if r[i_ind] is not None else "",
        }
    return out


# --- Category mapping -------------------------------------------------------
# SETSmart groups everything under Revenue / Expenses / Other Comprehensive Income
# by indent, but several rows are misleading:
#   - Net Investment Income, Other Gains (Losses) + subs, Increase (Decrease) In
#     Net Assets From Operations, and Total Comprehensive Income live under the
#     "Expenses" indent but are really subtotals / results, not expenses.
#   - "Increase (Decrease) In Net Assets From Operations" can appear twice (once
#     after Total Expenses, once at the start of the OCI section). For most REITs
#     the two values are identical, but for some (AIMIRT, DREIT, ...) they differ
#     due to restatements — so we keep both as distinct entries in different
#     categories.

PROFIT_RESULT_NAMES = {
    "Net Investment Income",
    "Other Gains (Losses)",
    "Increase (Decrease) In Net Assets From Operations",
    "Total Comprehensive Income (Expense) For The Period",
}

OCI_PATH_PREFIXES = (
    "Other Comprehensive Income >",
    "Items That Will Be Subsequently Reclassified",
    "Items That Will Not Be Subsequently Reclassified",
)

CATEGORY_ORDER = ["Revenue", "Expenses", "Profit & Results", "Other Comprehensive Income"]

ITEM_ORDER: dict[str, list[str]] = {
    "Revenue": [
        "Total Revenue",
        "Revenue From Investments",
        "Interest Income",
        "Rental And Service Income",
        "Income From Guarantee Of Profit And/Or Net Income",
        "Other Income",
    ],
    "Expenses": [
        "Total Expenses",
        "Costs Of Rental And Services",
        "Finance Costs",
        "Property Management Fee",
        "Management Fee",
        "Trustee Fee",
        "Selling And Administrative Expenses",
        "Selling Expenses",
        "Administrative Expenses",
        "Income Tax Expense",
        "Professional Fees",
        "Audit Fee",
        "Professional Fees - Others",
        "Registrar Fee",
        "Common Area Management Fee",
        "(Reversal Of) Expected Credit Losses",
        "(Reversal Of) Loss On Impairment",
        "Deferred Expense Amortisation",
        "Membership Fee",
        "Other Expenses",
    ],
    "Profit & Results": [
        "Net Investment Income",
        "Other Gains (Losses)",
        "Gains (Losses) From Investments",
        "Gains (Losses) On Fair Value Adjustments Of Investments",
        "Gains (Losses) On Foreign Currency Exchange",
        "Gains (Losses) On Net Monetary Position",
        "Other Gains (Losses) - Others",
        "Increase (Decrease) In Net Assets From Operations",
        "Total Comprehensive Income (Expense) For The Period",
    ],
    "Other Comprehensive Income": [
        "Increase (Decrease) In Net Assets From Operations",
        "Other Comprehensive Income (Expense) - Net Of Tax",
        "Currency Translation Adjustments",
        "Remeasurement Of Employee Benefit Obligations",
    ],
}


def categorize(name: str, path: str) -> str:
    """Assign a line item to one of four logical categories.

    Order of rules matters — early rules short-circuit.
    """
    # 1. Revenue path is unambiguous
    if path.startswith("Revenue >"):
        return "Revenue"
    # 2. The "Increase (Decrease)..." line re-printed under the OCI section: keep
    #    it distinct from the canonical Profit & Results entry. Values can differ
    #    after restatements (AIMIRT, DREIT) so we don't dedupe them.
    if path.startswith("Other Comprehensive Income >"):
        return "Other Comprehensive Income"
    # 3. Result subtotals (sitting under the Expenses indent but not expenses)
    if name in PROFIT_RESULT_NAMES:
        return "Profit & Results"
    # 4. Any sub-item of "Other Gains (Losses)" is also a result component
    if "Other Gains (Losses)" in path:
        return "Profit & Results"
    # 5. Items inside "Items That Will [Not] Be Subsequently Reclassified" sections
    #    are OCI line items (currency translation, employee benefit remeasurement,
    #    OCI net-of-tax). EXCEPT the bottom-line "Total Comprehensive Income" which
    #    was already routed to Profit & Results by rule 3.
    if any(path.startswith(p) for p in OCI_PATH_PREFIXES):
        return "Other Comprehensive Income"
    return "Expenses"


def build_payload(rows: list[dict], labels: dict[str, dict] | None = None) -> dict:
    labels = labels or {}
    symbols = sorted({r["symbol"] for r in rows})
    periods = sorted({r["period"] for r in rows})  # calendar — 2022Q2 etc.
    period_end_by_period = {r["period"]: r["period_end"] for r in rows}

    # Coverage: how many symbols report each item_path
    coverage: dict[str, set] = defaultdict(set)
    item_meta: dict[str, dict] = {}
    for r in rows:
        path = r["item_path"]
        coverage[path].add(r["symbol"])
        if path not in item_meta:
            item_meta[path] = {"name": r["item"], "level": int(r["level"])}

    items = []
    for path in sorted(item_meta):
        name = item_meta[path]["name"]
        category = categorize(name, path)
        order_list = ITEM_ORDER.get(category, [])
        try:
            within_rank = order_list.index(name)
        except ValueError:
            within_rank = len(order_list)  # unmapped → push to end of category
        items.append({
            "path": path,
            "name": name,
            "level": item_meta[path]["level"],
            "coverage": len(coverage[path]),
            "category": category,
            "_cat_rank": CATEGORY_ORDER.index(category),
            "_within_rank": within_rank,
        })

    # De-duplicate strictly within (category, name) — same leaf-name appearing
    # in two different categories is intentionally KEPT so cross-category
    # duplicates (e.g. "Increase (Decrease)..." in both Profit & Results and
    # OCI) remain visible, since their values can differ for restated stocks.
    seen: dict[tuple[str, str], dict] = {}
    for it in items:
        key = (it["category"], it["name"])
        if key not in seen or it["coverage"] > seen[key]["coverage"]:
            seen[key] = it
    items = list(seen.values())

    # Disambiguate display label: when the same leaf name shows up in multiple
    # categories, append " · <category>" so the dropdown is unambiguous.
    from collections import Counter as _Ctr
    name_count = _Ctr(it["name"] for it in items)
    cat_short = {
        "Revenue": "Revenue",
        "Expenses": "Expenses",
        "Profit & Results": "P&L",
        "Other Comprehensive Income": "OCI",
    }
    for it in items:
        it["display_name"] = (
            f"{it['name']}  ·  {cat_short.get(it['category'], it['category'])}"
            if name_count[it["name"]] > 1 else it["name"]
        )

    # Sort: by category order, then by manual within-category order, then by
    # coverage desc (for items not in the manual list), then path.
    items.sort(key=lambda x: (x["_cat_rank"], x["_within_rank"], -x["coverage"], x["path"]))
    for it in items:
        it.pop("_cat_rank")
        it.pop("_within_rank")

    # Nested values: symbol -> item_path -> period -> value
    values: dict[str, dict[str, dict[str, float]]] = defaultdict(lambda: defaultdict(dict))
    # Fiscal-period mapping: symbol -> calendar_period -> fiscal_period
    fiscal: dict[str, dict[str, str]] = defaultdict(dict)
    # Detect fiscal-year offset per symbol (quarters fiscal leads calendar)
    sym_offset: dict[str, int] = {}
    for r in rows:
        values[r["symbol"]][r["item_path"]][r["period"]] = float(r["value"])
        fiscal[r["symbol"]][r["period"]] = r["fiscal_period"]
        if r["symbol"] not in sym_offset:
            fy, fq = r["fiscal_period"].split("Q")
            cy, cq = r["period"].split("Q")
            sym_offset[r["symbol"]] = (int(fy) * 4 + int(fq)) - (int(cy) * 4 + int(cq))

    fiscal_year_starts = {1: "Oct", 2: "Jul", 3: "Apr"}
    non_calendar = [
        {
            "symbol": s,
            "offset": sym_offset[s],
            "fy_starts": fiscal_year_starts.get(sym_offset[s], "?"),
        }
        for s in sorted(sym_offset)
        if sym_offset[s] != 0
    ]

    # Labels: only include entries for symbols actually present in the data
    labels_clean = {s: labels[s] for s in symbols if s in labels}
    types = sorted({v["type"] for v in labels_clean.values() if v.get("type")})
    industries = sorted({v["industry"] for v in labels_clean.values() if v.get("industry")})

    # Warn for symbols missing a label so the analyst notices
    unlabeled = [s for s in symbols if s not in labels_clean]
    if unlabeled:
        print(f"  Symbols without labels: {', '.join(unlabeled)}")

    return {
        "meta": {
            "generated": datetime.now().isoformat(timespec="seconds"),
            "unit": "k.Baht",
            "row_count": len(rows),
            "non_calendar": non_calendar,
        },
        "symbols": symbols,
        "periods": periods,
        "period_ends": period_end_by_period,
        "items": items,
        "values": values,
        "fiscal": fiscal,
        "labels": labels_clean,
        "types": types,
        "industries": industries,
    }


HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>REIT & Property Fund Analytics</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root {
    --bg: #fafafa;
    --surface: #ffffff;
    --surface-2: #f7f7f8;
    --surface-hover: #f9fafb;
    --border: #e5e7eb;
    --border-strong: #d1d5db;
    --text: #111827;
    --text-2: #374151;
    --muted: #6b7280;
    --muted-2: #9ca3af;
    --accent: #2563eb;
    --accent-soft: #eff6ff;
    --accent-text: #1e40af;
    --pos: #16a34a;
    --neg: #dc2626;
    --pos-bg: #f0fdf4;
    --neg-bg: #fef2f2;
    --shadow-sm: 0 1px 2px rgba(16,24,40,.04);
    --radius: 6px;
    --radius-lg: 8px;
  }
  * { box-sizing: border-box; }
  html, body {
    margin: 0;
    background: var(--bg);
    color: var(--text);
    font-family: "Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, sans-serif;
    font-size: 13px;
    line-height: 1.45;
    font-feature-settings: 'cv11', 'ss01', 'ss03';
    -webkit-font-smoothing: antialiased;
    -moz-osx-font-smoothing: grayscale;
  }
  ::selection { background: #dbeafe; }

  /* ============ LAYOUT: sidebar + main ============ */
  .app { display: grid; grid-template-columns: 236px 1fr; min-height: 100vh; }
  /* Mobile-only topbar + drawer overlay — hidden on desktop */
  .topbar { display: none; }
  .sidebar-overlay { display: none; }
  .sidebar {
    background: var(--surface-2);
    border-right: 1px solid var(--border);
    padding: 20px 14px;
    display: flex; flex-direction: column; gap: 18px;
    position: sticky; top: 0; height: 100vh;
  }
  .brand { padding: 2px 8px 6px; }
  .brand .title { font-size: 14px; font-weight: 600; color: var(--text); letter-spacing: -0.01em; }
  .brand .subtitle { font-size: 11px; color: var(--muted); margin-top: 2px; }

  .nav-list { display: flex; flex-direction: column; gap: 2px; }
  .nav-label {
    font-size: 10px; font-weight: 600; color: var(--muted-2);
    text-transform: uppercase; letter-spacing: 0.06em;
    padding: 0 8px 6px;
  }
  #tabs { display: flex; flex-direction: column; gap: 1px; }
  #tabs button {
    background: transparent; border: none; cursor: pointer;
    text-align: left; padding: 7px 10px;
    font: inherit; font-size: 13px; font-weight: 500;
    color: var(--text-2);
    border-radius: 6px;
    display: flex; align-items: center; gap: 9px;
    transition: background 80ms ease, color 80ms ease;
  }
  #tabs button:hover:not(.active) { background: rgba(17,24,39,.04); color: var(--text); }
  #tabs button.active {
    background: var(--surface);
    color: var(--text);
    box-shadow: inset 0 0 0 1px var(--border), var(--shadow-sm);
  }
  #tabs button .dot {
    width: 6px; height: 6px; border-radius: 50%;
    background: var(--muted-2);
    flex-shrink: 0;
  }
  #tabs button.active .dot { background: var(--accent); }

  .sidebar-meta {
    margin-top: auto;
    font-size: 11px; color: var(--muted); line-height: 1.5;
    padding: 10px 8px; border-top: 1px solid var(--border);
  }
  .sidebar-meta .row { display: flex; justify-content: space-between; gap: 8px; }
  .sidebar-meta .row + .row { margin-top: 3px; }
  .sidebar-meta .k { color: var(--muted-2); }
  .sidebar-meta .v { color: var(--text-2); font-variant-numeric: tabular-nums; }

  /* ============ MAIN CONTENT ============ */
  main {
    padding: 22px 28px 40px;
    max-width: 1600px;
    width: 100%;
  }
  .page-head {
    display: flex; align-items: baseline; justify-content: space-between;
    gap: 16px; margin-bottom: 16px;
  }
  .page-head h2 {
    margin: 0; font-size: 18px; font-weight: 600;
    color: var(--text); letter-spacing: -0.01em;
  }
  .page-head .page-sub { font-size: 12px; color: var(--muted); }

  section.tab { display: none; }
  section.tab.active { display: block; }

  /* ============ FY BANNER ============ */
  #fy-banner {
    display: flex; gap: 10px; align-items: flex-start;
    padding: 9px 12px; margin-bottom: 14px;
    background: var(--accent-soft);
    border: 1px solid #dbeafe;
    border-radius: var(--radius);
    color: var(--accent-text);
    font-size: 12px; line-height: 1.5;
  }
  #fy-banner .icon {
    flex-shrink: 0; width: 14px; height: 14px;
    margin-top: 1px;
    color: var(--accent);
  }
  #fy-banner b { font-weight: 600; color: #1e3a8a; }
  #fy-banner .syms { color: var(--accent-text); font-variant-numeric: tabular-nums; }

  /* ============ CONTROLS / FILTER BAR ============ */
  .controls {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 10px 14px; align-items: end;
    padding: 12px 14px;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius-lg);
    margin-bottom: 14px;
  }
  /* TS tab: 2-column layout. Left = symbols card (full height). Right = options + chart + summary stacked. */
  .ts-layout {
    display: grid;
    grid-template-columns: 260px 1fr;
    gap: 14px;
    align-items: stretch;
  }
  /* Card itself is a flex column so the select can grow with flex: 1.
     `display: contents` on the inner <label> lets its children participate
     directly in the card's flex layout — without it, native <select multiple>
     ignores flex-grow and falls back to its intrinsic 4-row height. */
  .ts-layout > .ts-symbols-card {
    display: flex;
    flex-direction: column;
    align-items: stretch;   /* override .controls' align-items: end, which would right-align the label & hint in this flex-column layout */
    align-self: stretch;
    margin-bottom: 0;
    padding: 14px;
    gap: 10px;
    min-height: 0;
    overflow: hidden;
  }
  /* Pair Fund type + Industry as a 2-up filter row so they read as a group
     attached to the Symbols list below. Flex (not grid) so each dropdown
     can size to its content and the pair sits flush against the left edge
     instead of stretching to fill the row. */
  .ts-symbols-card .filter-row {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    padding-bottom: 10px;
    border-bottom: 1px solid var(--border);
  }
  .ts-symbols-card .filter-row label {
    margin: 0;
    flex: 0 0 auto;
  }
  .ts-symbols-card .filter-row select {
    width: auto;
  }
  .ts-symbols-card label.full { display: contents; }
  .ts-symbols-card .label-text {
    font-size: 11px; font-weight: 500; color: var(--muted);
    letter-spacing: 0.01em;
  }
  .ts-symbols-card select[multiple] {
    flex: 1 1 auto;
    width: 100%;
    min-width: 0;
    min-height: 260px;     /* baseline so it never collapses below this */
    height: 100%;
    box-sizing: border-box;
    background: var(--surface);
    background-image: none;
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 4px;
    font-variant-numeric: tabular-nums;
    overflow-y: auto;
  }
  .ts-symbols-card select[multiple]:focus {
    border-color: var(--accent);
    box-shadow: 0 0 0 3px rgba(37,99,235,.12);
    outline: none;
  }
  .ts-symbols-card select[multiple] option {
    padding: 5px 8px;
    font-size: 12.5px;
    border-radius: 3px;
    line-height: 1.45;
  }
  .ts-symbols-card .hint {
    font-size: 11px; color: var(--muted-2); font-weight: 400;
    margin-top: 2px;
  }
  .ts-right {
    display: flex; flex-direction: column; gap: 14px;
    min-width: 0;       /* allow chart canvas to shrink */
  }
  .ts-right .controls,
  .ts-right .panel { margin-bottom: 0; }
  @media (max-width: 980px) {
    .ts-layout { grid-template-columns: 1fr; }
  }
  .controls label {
    display: flex; flex-direction: column; gap: 5px;
    font-size: 11px; font-weight: 500;
    color: var(--muted);
    min-width: 0;          /* allow children to shrink inside grid cells */
  }
  .controls select,
  .controls input[type=text] {
    background: var(--surface);
    color: var(--text);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 0 10px;
    height: 30px;
    font: inherit;
    font-size: 13px;
    width: 100%;            /* fill the grid cell instead of overflowing it */
    min-width: 0;           /* allow shrinking when cell is narrow */
    max-width: 100%;
    text-overflow: ellipsis; /* selected value truncates with … on narrow widths */
    overflow: hidden;
    white-space: nowrap;
    transition: border-color 80ms ease, box-shadow 80ms ease;
    appearance: none; -webkit-appearance: none;
  }
  .controls select {
    background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 12 12'><path fill='none' stroke='%236b7280' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round' d='M3 4.5l3 3 3-3'/></svg>");
    background-repeat: no-repeat;
    background-position: right 8px center;
    background-size: 12px;
    padding-right: 28px;
  }
  .controls select[multiple] {
    background-image: none;
    padding: 4px 6px;
    height: auto;
    min-height: 92px;
    font-variant-numeric: tabular-nums;
  }
  .controls select[multiple] option {
    padding: 3px 6px;
    border-radius: 3px;
    font-size: 12px;
  }
  .controls select optgroup {
    font-size: 10px;
    font-weight: 600;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: var(--muted-2);
    background: var(--surface-2);
  }
  .controls select option {
    font-weight: 400;
    color: var(--text);
    background: var(--surface);
    text-transform: none;
    letter-spacing: 0;
  }
  /* Category pill in raw-data table */
  .cat-pill {
    display: inline-block;
    padding: 1px 7px;
    font-size: 10px;
    font-weight: 500;
    letter-spacing: 0.02em;
    border-radius: 10px;
    border: 1px solid transparent;
    line-height: 1.5;
    white-space: nowrap;
  }
  .cat-Revenue          { background: #ecfdf5; color: #047857; border-color: #d1fae5; }
  .cat-Expenses         { background: #fef2f2; color: #b91c1c; border-color: #fee2e2; }
  .cat-ProfitResults    { background: #eff6ff; color: #1d4ed8; border-color: #dbeafe; }
  .cat-OtherComprehensiveIncome { background: #faf5ff; color: #7e22ce; border-color: #f3e8ff; }
  .cat-Other            { background: var(--surface-2); color: var(--muted); border-color: var(--border); }

  /* Industry + Type pills — Stripe-inspired soft palette */
  .ind-pill, .type-pill {
    display: inline-flex; align-items: center; gap: 5px;
    padding: 1px 8px;
    font-size: 10.5px;
    font-weight: 500;
    letter-spacing: 0.01em;
    border-radius: 10px;
    border: 1px solid transparent;
    line-height: 1.55;
    white-space: nowrap;
  }
  .ind-pill::before {
    content: ""; width: 6px; height: 6px; border-radius: 50%;
    background: currentColor; opacity: 0.85; flex-shrink: 0;
  }
  /* Industry palette (soft bg + dark text, indicator dot inherits text) */
  .ind-Retail            { background:#fdf2f8; color:#be185d; border-color:#fbcfe8; }
  .ind-Industrial        { background:#eef2ff; color:#4338ca; border-color:#c7d2fe; }
  .ind-Office            { background:#e0f2fe; color:#075985; border-color:#bae6fd; }
  .ind-Airport           { background:#fff7ed; color:#c2410c; border-color:#fed7aa; }
  .ind-Hospitality       { background:#fef3c7; color:#a16207; border-color:#fde68a; }
  .ind-Mixed             { background:#f5f3ff; color:#6d28d9; border-color:#ddd6fe; }
  .ind-DataCenter        { background:#ecfdf5; color:#047857; border-color:#a7f3d0; }
  .ind-ConventionCenter  { background:#fae8ff; color:#a21caf; border-color:#f5d0fe; }
  .ind-Residential       { background:#ccfbf1; color:#0f766e; border-color:#99f6e4; }
  .ind-Selfstorage       { background:#f1f5f9; color:#334155; border-color:#cbd5e1; }
  .ind-Unknown           { background:var(--surface-2); color:var(--muted); border-color:var(--border); }

  /* Type pills */
  .type-REIT          { background:#eef2ff; color:#4f46e5; border-color:#e0e7ff; }
  .type-PropertyFund  { background:#f1f5f9; color:#475569; border-color:#cbd5e1; }
  .type-Unknown       { background:var(--surface-2); color:var(--muted); border-color:var(--border); }

  .controls input[type=text]::placeholder { color: var(--muted-2); }
  .controls select:hover,
  .controls input[type=text]:hover { border-color: var(--border-strong); }
  .controls select:focus,
  .controls input[type=text]:focus {
    outline: none;
    border-color: var(--accent);
    box-shadow: 0 0 0 3px rgba(37,99,235,.12);
  }


  /* ============ PANELS ============ */
  .panel {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius-lg);
    padding: 14px 16px;
  }
  .panel + .panel { margin-top: 14px; }
  .panel-head {
    display: flex; align-items: baseline; justify-content: space-between;
    gap: 12px; margin-bottom: 10px;
  }
  .panel-head h3 {
    margin: 0; font-size: 13px; font-weight: 600; color: var(--text);
    letter-spacing: -0.005em;
  }
  .panel-head .sub { font-size: 12px; color: var(--muted); }

  .chart-wrap { position: relative; height: 420px; }
  .chart-wrap.tall { height: 620px; }

  /* ============ TABLES ============ */
  table.data {
    width: 100%; border-collapse: collapse;
    font-size: 12.5px;
    font-variant-numeric: tabular-nums;
  }
  table.data th, table.data td {
    padding: 7px 10px;
    border-bottom: 1px solid var(--border);
    text-align: right;
    white-space: nowrap;
    line-height: 1.4;
  }
  table.data thead th {
    background: var(--surface);
    position: sticky; top: 0; z-index: 1;
    cursor: pointer; user-select: none;
    font-size: 11.5px;
    font-weight: 500;
    color: var(--muted);
    border-bottom: 1px solid var(--border-strong);
  }
  table.data thead th:hover { color: var(--text); }
  table.data th.sort-asc, table.data th.sort-desc { color: var(--accent); }
  table.data th.sort-asc::after  { content: " ▲"; font-size: 9px; color: var(--accent); }
  table.data th.sort-desc::after { content: " ▼"; font-size: 9px; color: var(--accent); }
  table.data td.sym, table.data th.sym {
    text-align: left; font-weight: 500; color: var(--text);
  }
  table.data td.item, table.data th.item { text-align: left; color: var(--text-2); }
  table.data tbody tr:hover td { background: var(--surface-hover); }
  table.data tbody tr:last-child td { border-bottom: none; }

  /* ============ HEATMAP TABLE ============ */
  table.heat {
    width: 100%; border-collapse: separate; border-spacing: 0;
    font-size: 11.5px;
    font-variant-numeric: tabular-nums;
  }
  table.heat th, table.heat td {
    padding: 5px 8px;
    text-align: right;
    border-right: 1px solid var(--border);
    border-bottom: 1px solid var(--border);
    white-space: nowrap;
    line-height: 1.3;
  }
  table.heat thead th {
    background: var(--surface);
    color: var(--muted);
    font-weight: 500;
    font-size: 11px;
    position: sticky; top: 0; z-index: 3;
    border-bottom: 1px solid var(--border-strong);
    cursor: pointer; user-select: none;
  }
  table.heat thead th:hover { color: var(--text); }
  table.heat th.sort-asc, table.heat th.sort-desc { color: var(--accent); }
  table.heat th.sort-asc::after  { content: " ▲"; font-size: 9px; color: var(--accent); }
  table.heat th.sort-desc::after { content: " ▼"; font-size: 9px; color: var(--accent); }
  table.heat th.sym {
    text-align: left;
    position: sticky; left: 0; z-index: 4;
    background: var(--surface);
  }
  table.heat td.sym {
    text-align: left;
    font-weight: 500;
    color: var(--text);
    background: var(--surface);
    position: sticky; left: 0; z-index: 2;
    border-right: 1px solid var(--border-strong);
  }
  table.heat td.empty {
    color: var(--muted-2);
    background:
      repeating-linear-gradient(135deg,
        #fafafa 0, #fafafa 4px,
        #f3f4f6 4px, #f3f4f6 8px) !important;
  }
  table.heat tbody tr:hover td.sym { background: var(--surface-hover); }

  /* ============ MISC ============ */
  .pos { color: var(--pos); }
  .neg { color: var(--neg); }
  .scroll { overflow: auto; max-height: 72vh; }
  .h-scroll { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  .scroll table.data, .scroll table.heat { font-variant-numeric: tabular-nums; }
  /* When a table is inside a .scroll wrapper, drop its own borders */
  .scroll table.data thead th { border-top: none; }
  /* Flush the scroll-table inside a panel to the panel edges */
  .panel.flush { padding: 0; }
  .panel.flush > .scroll { border-radius: var(--radius-lg); max-height: 72vh; }
  .panel.flush table.data thead th:first-child,
  .panel.flush table.heat thead th:first-child,
  .panel.flush table.data tbody td:first-child,
  .panel.flush table.heat tbody td:first-child { padding-left: 16px; }
  .panel.flush table.data thead th:last-child,
  .panel.flush table.data tbody td:last-child { padding-right: 16px; }
  .panel.flush .footnote { padding: 10px 16px 12px; margin: 0; border-top: 1px solid var(--border); }

  .footnote {
    font-size: 11.5px; color: var(--muted);
    margin-top: 10px; line-height: 1.5;
  }
  .footnote b { color: var(--text-2); font-weight: 600; }
  .row-count {
    font-size: 11.5px; color: var(--muted);
    font-variant-numeric: tabular-nums;
    height: 30px; display: flex; align-items: center;
  }
  .fp-note {
    color: var(--muted); font-size: 11px;
    margin-left: 6px; font-weight: 400;
    font-variant-numeric: tabular-nums;
  }
  .empty-state {
    color: var(--muted); font-size: 12px;
    padding: 4px 0; font-style: normal;
  }

  /* Scrollbar polish (webkit) */
  ::-webkit-scrollbar { width: 10px; height: 10px; }
  ::-webkit-scrollbar-thumb {
    background: #d1d5db; border-radius: 5px;
    border: 2px solid transparent; background-clip: padding-box;
  }
  ::-webkit-scrollbar-thumb:hover { background: #9ca3af; background-clip: padding-box; border: 2px solid transparent; }
  ::-webkit-scrollbar-track { background: transparent; }

  /* ============ RESPONSIVE / MOBILE ============ */
  /* Tablet/phone: sidebar becomes an off-canvas drawer, opened via hamburger */
  @media (max-width: 900px) {
    .app { grid-template-columns: 1fr; grid-template-rows: auto 1fr; }
    .topbar {
      display: flex; align-items: center; gap: 12px;
      grid-column: 1; grid-row: 1;
      position: sticky; top: 0; z-index: 100;
      background: var(--surface);
      border-bottom: 1px solid var(--border);
      padding: 10px 14px;
    }
    .topbar-title {
      font-size: 14px; font-weight: 600;
      color: var(--text); letter-spacing: -0.01em;
    }
    .hamburger {
      width: 34px; height: 34px;
      background: transparent;
      border: 1px solid var(--border);
      border-radius: 6px;
      cursor: pointer; padding: 0;
      display: flex; flex-direction: column;
      justify-content: center; align-items: center;
      gap: 4px;
      transition: background 80ms ease, border-color 80ms ease;
    }
    .hamburger:hover { background: var(--surface-2); border-color: var(--border-strong); }
    .hamburger:active { background: rgba(17,24,39,.06); }
    .hamburger span {
      display: block; width: 16px; height: 1.6px;
      background: var(--text-2); border-radius: 1px;
    }
    .sidebar {
      position: fixed;
      top: 0; left: 0; bottom: 0;
      width: 260px; max-width: 82vw;
      height: 100vh;
      z-index: 200;
      transform: translateX(-100%);
      transition: transform 220ms ease;
      border-right: 1px solid var(--border);
      overflow-y: auto;
    }
    .sidebar.open { transform: translateX(0); box-shadow: 4px 0 24px rgba(16,24,40,.12); }
    .sidebar-overlay {
      display: block;
      position: fixed; inset: 0;
      background: rgba(0,0,0,0);
      z-index: 150;
      pointer-events: none;
      transition: background 220ms ease;
    }
    .sidebar-overlay.open {
      background: rgba(15,23,42,.45);
      pointer-events: auto;
    }
    main { padding: 16px; max-width: 100%; grid-column: 1; grid-row: 2; }
    .ts-layout { grid-template-columns: 1fr; }
    /* In single-column layout the symbols card must not flex-stretch — let the
       <select multiple> take its natural height (compact button on iOS Safari,
       a fixed-size listbox elsewhere) instead of inflating to fill the row. */
    .ts-layout > .ts-symbols-card { overflow: visible; }
    .ts-symbols-card select[multiple] {
      flex: 0 0 auto;
      height: auto;
      min-height: 140px;
    }
    .chart-wrap { height: 360px; }
    .chart-wrap.tall { height: 520px; }
  }

  /* Phone: stack controls vertically, shrink everything */
  @media (max-width: 640px) {
    html, body { font-size: 12.5px; }
    main { padding: 12px 10px 28px; }
    .page-head { flex-direction: column; align-items: flex-start; gap: 4px; margin-bottom: 12px; }
    .page-head h2 { font-size: 16px; }
    .page-head .page-sub { font-size: 11.5px; }
    #fy-banner { font-size: 11.5px; padding: 8px 10px; }

    .controls {
      grid-template-columns: 1fr;
      gap: 8px;
      padding: 10px 12px;
    }
    .controls label { font-size: 11px; }
    .controls select, .controls input[type=text] { height: 32px; font-size: 13px; }

    .panel { padding: 10px 12px; }
    .panel.flush > .scroll { max-height: 60vh; }
    .chart-wrap { height: 280px; }
    .chart-wrap.tall { height: 440px; }
    .scroll { max-height: 60vh; }

    /* Tables: shrink horizontal padding so more fits per row */
    table.data th, table.data td { padding: 6px 6px; font-size: 11.5px; }
    table.heat th, table.heat td { padding: 4px 6px; font-size: 11px; }
    .panel.flush table.data thead th:first-child,
    .panel.flush table.heat thead th:first-child,
    .panel.flush table.data tbody td:first-child,
    .panel.flush table.heat tbody td:first-child { padding-left: 10px; }
    .panel.flush table.data thead th:last-child,
    .panel.flush table.data tbody td:last-child { padding-right: 10px; }

    /* TS symbols card: trim padding so symbols list isn't tiny */
    .ts-layout > .ts-symbols-card { padding: 10px; gap: 6px; }
    .ts-symbols-card select[multiple] { min-height: 120px; }

    /* Sidebar brand: hide the subtitle on phones to save room */
    .sidebar .brand .subtitle { display: none; }
  }

  /* Very narrow phones */
  @media (max-width: 380px) {
    #tabs button { padding: 6px 8px; font-size: 11.5px; }
    #tabs button .dot { width: 5px; height: 5px; }
    .sidebar .brand .title { font-size: 13px; }
  }
</style>
</head>
<body>
<div class="app">
  <header class="topbar">
    <button class="hamburger" id="hamburger" aria-label="Open navigation" aria-expanded="false">
      <span></span><span></span><span></span>
    </button>
    <div class="topbar-title">REIT Analytics</div>
  </header>
  <div class="sidebar-overlay" id="sidebar-overlay" aria-hidden="true"></div>
  <aside class="sidebar">
    <div class="brand">
      <div class="title">REIT Analytics</div>
      <div class="subtitle">Thai market &middot; quarterly</div>
    </div>
    <div class="nav-list">
      <div class="nav-label">Views</div>
      <div id="tabs">
        <button data-tab="timeseries" class="active"><span class="dot"></span>Time series</button>
        <button data-tab="peer"><span class="dot"></span>Peer comparison</button>
        <button data-tab="heatmap"><span class="dot"></span>QoQ / YoY heatmap</button>
        <button data-tab="raw"><span class="dot"></span>Raw data</button>
      </div>
    </div>
    <div class="sidebar-meta">
      <div class="row"><span class="k">Symbols</span><span class="v">__SYMBOL_COUNT__</span></div>
      <div class="row"><span class="k">Quarters</span><span class="v">__PERIOD_COUNT__</span></div>
      <div class="row"><span class="k">Data points</span><span class="v">__ROW_COUNT__</span></div>
      <div class="row"><span class="k">Unit</span><span class="v">k.Baht</span></div>
      <div class="row"><span class="k">Generated</span><span class="v">__GENERATED__</span></div>
    </div>
  </aside>

<main>
<div id="fy-banner"></div>

<!-- ========== TIME SERIES ========== -->
<section id="tab-timeseries" class="tab active">
  <div class="page-head">
    <h2>Time series</h2>
    <div class="page-sub">Quarterly trend for one or more REITs, side by side.</div>
  </div>
  <div class="ts-layout">
    <div class="controls ts-symbols-card">
      <div class="filter-row">
        <label>Fund type
          <select id="ts-type"><option value="">(All)</option></select>
        </label>
        <label>Industry
          <select id="ts-industry"><option value="">(All)</option></select>
        </label>
      </div>
      <label class="full">
        <span class="label-text">Symbols</span>
        <select id="ts-symbols" multiple></select>
        <span class="hint" id="ts-symbols-hint">Cmd/Ctrl-click to pick several</span>
      </label>
    </div>
    <div class="ts-right">
      <div class="controls ts-options-card">
        <label>Line item
          <select id="ts-item"></select>
        </label>
        <label>Mode
          <select id="ts-mode">
            <option value="abs">Absolute</option>
            <option value="qoq">QoQ %</option>
            <option value="yoy">YoY %</option>
            <option value="idx">Index = 100</option>
          </select>
        </label>
        <label>Range
          <select id="ts-range">
            <option value="all">All periods</option>
            <option value="3y">Last 3 years</option>
            <option value="2y">Last 2 years</option>
            <option value="1y">Last 1 year</option>
          </select>
        </label>
      </div>
      <div class="panel">
        <div class="chart-wrap"><canvas id="ts-chart"></canvas></div>
      </div>
      <div class="panel">
        <div id="ts-summary"></div>
      </div>
    </div>
  </div>
</section>

<!-- ========== PEER COMPARISON ========== -->
<section id="tab-peer" class="tab">
  <div class="page-head">
    <h2>Peer comparison</h2>
    <div class="page-sub">Rank all REITs on a single line item for a chosen quarter.</div>
  </div>
  <div class="controls">
    <label>Line item
      <select id="peer-item"></select>
    </label>
    <label>Quarter
      <select id="peer-period"></select>
    </label>
    <label>Fund type
      <select id="peer-type"><option value="">(All)</option></select>
    </label>
    <label>Industry
      <select id="peer-industry"><option value="">(All)</option></select>
    </label>
    <label>Sort
      <select id="peer-sort">
        <option value="desc">High → Low</option>
        <option value="asc">Low → High</option>
        <option value="alpha">A → Z</option>
      </select>
    </label>
    <label>Symbol filter
      <input type="text" id="peer-filter" placeholder="contains…">
    </label>
    <label>Compare against
      <select id="peer-compare">
        <option value="abs">Absolute</option>
        <option value="qoq">vs. prior Q (%)</option>
        <option value="yoy">YoY (%)</option>
        <option value="rev">% of Total Revenue</option>
      </select>
    </label>
  </div>
  <div class="panel">
    <div class="chart-wrap tall"><canvas id="peer-chart"></canvas></div>
  </div>
  <div class="panel flush">
    <div class="scroll">
      <table class="data" id="peer-table"></table>
    </div>
  </div>
</section>

<!-- ========== HEATMAP ========== -->
<section id="tab-heatmap" class="tab">
  <div class="page-head">
    <h2>QoQ / YoY heatmap</h2>
    <div class="page-sub">Symbol &times; quarter matrix, colour-scaled by change.</div>
  </div>
  <div class="controls">
    <label>Line item
      <select id="heat-item"></select>
    </label>
    <label>Display
      <select id="heat-mode">
        <option value="qoq">QoQ %</option>
        <option value="yoy">YoY %</option>
        <option value="abs">Absolute</option>
      </select>
    </label>
    <label>Fund type
      <select id="heat-type"><option value="">(All)</option></select>
    </label>
    <label>Industry
      <select id="heat-industry"><option value="">(All)</option></select>
    </label>
    <label>Symbol filter
      <input type="text" id="heat-filter" placeholder="contains…">
    </label>
    <label>Sort by
      <select id="heat-sort">
        <option value="alpha">A → Z</option>
        <option value="last">Latest value</option>
        <option value="latestqoq">Latest QoQ %</option>
      </select>
    </label>
  </div>
  <div class="panel flush">
    <div class="scroll" style="max-height:75vh"><table class="heat" id="heat-table"></table></div>
    <div class="footnote">Green = positive change &middot; Red = negative change &middot; Hatched cell = no data filed for that quarter.</div>
  </div>
</section>

<!-- ========== RAW DATA ========== -->
<section id="tab-raw" class="tab">
  <div class="page-head">
    <h2>Raw data</h2>
    <div class="page-sub">Filter and sort every line item in long format.</div>
  </div>
  <div class="controls">
    <label>Symbol contains
      <input type="text" id="raw-sym" placeholder="e.g. CPN">
    </label>
    <label>Fund type
      <select id="raw-type"><option value="">(all)</option></select>
    </label>
    <label>Industry
      <select id="raw-industry"><option value="">(all)</option></select>
    </label>
    <label>Category
      <select id="raw-cat">
        <option value="">(all)</option>
        <option value="Revenue">Revenue</option>
        <option value="Expenses">Expenses</option>
        <option value="Profit &amp; Results">Profit &amp; Results</option>
        <option value="Other Comprehensive Income">Other Comprehensive Income</option>
      </select>
    </label>
    <label>Item contains
      <input type="text" id="raw-item" placeholder="e.g. Revenue">
    </label>
    <label>Period
      <select id="raw-period"><option value="">(any)</option></select>
    </label>
    <label>Min |value|
      <input type="text" id="raw-min" placeholder="e.g. 1000">
    </label>
    <label>&nbsp;<span class="row-count" id="raw-count"></span></label>
  </div>
  <div class="panel flush">
    <div class="scroll"><table class="data" id="raw-table"></table></div>
  </div>
</section>

</main>
</div><!-- /.app -->

<script id="data" type="application/json">__DATA_JSON__</script>
<script>
"use strict";
const DATA = JSON.parse(document.getElementById('data').textContent);

// Render fiscal-year banner
(function renderFyBanner(){
  const nc = DATA.meta.non_calendar || [];
  const el = document.getElementById('fy-banner');
  if (!nc.length){ el.style.display = 'none'; return; }
  const groups = {};
  for (const x of nc) (groups[x.fy_starts] = groups[x.fy_starts] || []).push(x.symbol);
  const parts = Object.entries(groups).map(([start,syms]) =>
    `<b>FY starts ${start}:</b> <span class="syms">${syms.join(', ')}</span>`);
  const icon = `<svg class="icon" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">`
             + `<circle cx="8" cy="8" r="7" stroke="currentColor" stroke-width="1.4"/>`
             + `<path d="M8 7.25v3.5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>`
             + `<circle cx="8" cy="5.25" r="0.85" fill="currentColor"/></svg>`;
  el.innerHTML = icon + `<div>All periods aligned to <b>calendar quarters</b> (Jan–Mar = Q1, etc.). `
               + `${nc.length} symbols use non-calendar fiscal years — `
               + parts.join(' &middot; ')
               + `. Hover any value to see the original fiscal label.</div>`;
})();
function fiscalOf(sym, period){
  return (DATA.fiscal && DATA.fiscal[sym] && DATA.fiscal[sym][period]) || null;
}

// ----- helpers -----
const fmt = (v) => v == null || isNaN(v) ? "" :
  Math.abs(v) >= 1e6 ? (v/1e3).toLocaleString(undefined,{maximumFractionDigits:0})+" M"
  : v.toLocaleString(undefined,{maximumFractionDigits:0});
const fmtFull = (v) => v == null || isNaN(v) ? "—" : Math.round(v).toLocaleString(undefined,{maximumFractionDigits:0});
const pct = (v) => v == null || !isFinite(v) ? "—" :
  (v >= 0 ? "+" : "") + (v*100).toFixed(1) + "%";
const periodSort = (a,b) => a.localeCompare(b); // 2023Q1 sorts naturally
const prevQ = (p) => {
  const [y,q] = p.split('Q').map(Number);
  return q === 1 ? `${y-1}Q4` : `${y}Q${q-1}`;
};
const yoyQ = (p) => { const [y,q] = p.split('Q').map(Number); return `${y-1}Q${q}`; };
const periodLabel = (p) => p;  // already compact
const palette = ["#2563eb","#16a34a","#dc2626","#9333ea","#0891b2","#ea580c","#65a30d","#db2777","#7c3aed","#0d9488","#facc15","#475569"];
// Shared Chart.js theme tokens for the light dashboard
const CHART_FONT = '"Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, sans-serif';
const CHART_THEME = {
  tickColor: '#6b7280',
  gridColor: '#f3f4f6',
  legendColor: '#374151',
  borderColor: '#e5e7eb',
  axisTitleColor: '#6b7280',
};
const TOOLTIP_STYLE = {
  backgroundColor: '#ffffff',
  titleColor: '#111827',
  bodyColor: '#374151',
  borderColor: '#e5e7eb',
  borderWidth: 1,
  padding: 10,
  titleFont: { family: CHART_FONT, weight: '600', size: 12 },
  bodyFont:  { family: CHART_FONT, weight: '400', size: 12 },
  cornerRadius: 6,
  displayColors: true,
  boxPadding: 4,
};
if (window.Chart) {
  Chart.defaults.font.family = CHART_FONT;
  Chart.defaults.font.size = 12;
  Chart.defaults.color = CHART_THEME.legendColor;
  Chart.defaults.borderColor = CHART_THEME.gridColor;
}

// Index helpers
const ITEMS_BY_PATH = Object.fromEntries(DATA.items.map(it => [it.path, it]));

// ----- label (type/industry) helpers -----
const LABELS = DATA.labels || {};
const TYPES = DATA.types || [];
const INDUSTRIES = DATA.industries || [];
// Saturated palette used to colour chart bars by industry.
const INDUSTRY_COLORS = {
  'Retail':            '#ec4899',
  'Industrial':        '#6366f1',
  'Office':            '#0ea5e9',
  'Airport':           '#f97316',
  'Hospitality':       '#f59e0b',
  'Mixed':             '#8b5cf6',
  'Data Center':       '#10b981',
  'Convention Center': '#d946ef',
  'Residential':       '#14b8a6',
  'Self-storage':      '#64748b',
};
function labelOf(sym){ return LABELS[sym] || { type:'', industry:'' }; }
function slug(s){ return (s || '').replace(/[^a-zA-Z]/g, ''); }
function industryClass(ind){ return ind ? `ind-${slug(ind)}` : 'ind-Unknown'; }
function typeClass(t){ return t ? `type-${slug(t)}` : 'type-Unknown'; }
function industryColor(ind){ return INDUSTRY_COLORS[ind] || '#94a3b8'; }
function industryPill(ind){
  if (!ind) return '<span class="ind-pill ind-Unknown">—</span>';
  return `<span class="ind-pill ${industryClass(ind)}">${ind}</span>`;
}
function typePill(t){
  if (!t) return '<span class="type-pill type-Unknown">—</span>';
  return `<span class="type-pill ${typeClass(t)}">${t}</span>`;
}
function symbolMatchesLabels(sym, typeF, industryF){
  const l = labelOf(sym);
  if (typeF && l.type !== typeF) return false;
  if (industryF && l.industry !== industryF) return false;
  return true;
}
function populateLabelSelect(sel, options){
  // preserve the leading "(All)" / "(any)" option that's already in the HTML
  const first = sel.querySelector('option');
  sel.innerHTML = '';
  if (first) sel.appendChild(first);
  for (const v of options){
    const o = document.createElement('option');
    o.value = v; o.textContent = v;
    sel.appendChild(o);
  }
}
function getSeries(sym, itemPath){
  const m = DATA.values[sym] && DATA.values[sym][itemPath];
  return m || {};
}
function periodsForSymbol(sym){
  // periods where ANY value exists
  const seen = new Set();
  const v = DATA.values[sym] || {};
  for (const ip in v) for (const p in v[ip]) seen.add(p);
  return [...seen].sort(periodSort);
}

function populateItemSelect(sel, opts){
  opts = opts || {};
  sel.innerHTML = "";
  // Group items by category and render as <optgroup>. DATA.items is already
  // sorted by (category order, within-category order, coverage desc).
  const groups = {};
  const order = [];
  for (const it of DATA.items){
    if (opts.minCoverage && it.coverage < opts.minCoverage) continue;
    const cat = it.category || "Other";
    if (!(cat in groups)){ groups[cat] = []; order.push(cat); }
    groups[cat].push(it);
  }
  for (const cat of order){
    const og = document.createElement('optgroup');
    og.label = cat;
    for (const it of groups[cat]){
      const o = document.createElement('option');
      o.value = it.path;
      const cov = opts.showCoverage === false ? "" : `  [${it.coverage}/${DATA.symbols.length}]`;
      // display_name is disambiguated for leaf-name collisions across categories
      o.textContent = `${it.display_name || it.name}${cov}`;
      o.title = it.path;
      og.appendChild(o);
    }
    sel.appendChild(og);
  }
}
function populateSymbolSelect(sel){
  sel.innerHTML = "";
  for (const s of DATA.symbols){
    const o = document.createElement('option'); o.value = s; o.textContent = s;
    sel.appendChild(o);
  }
}
function populatePeriodSelect(sel, includeAny){
  sel.innerHTML = includeAny ? '<option value="">(any)</option>' : "";
  for (const p of [...DATA.periods].sort(periodSort).reverse()){
    const o = document.createElement('option'); o.value = p; o.textContent = p;
    sel.appendChild(o);
  }
}
function wireSelect(id, onChange){
  const el = document.getElementById(id);
  el.addEventListener('change', () => onChange(el.value));
}

// ----- tab switching -----
document.querySelectorAll('#tabs button').forEach(b => {
  b.addEventListener('click', () => {
    document.querySelectorAll('#tabs button').forEach(x => x.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
    b.classList.add('active');
    document.getElementById('tab-' + b.dataset.tab).classList.add('active');
    // On mobile, close the drawer after a tab is picked
    closeSidebar();
  });
});

// ----- mobile drawer (hamburger menu) -----
const _sidebar = document.querySelector('.sidebar');
const _overlay = document.getElementById('sidebar-overlay');
const _hamburger = document.getElementById('hamburger');
function openSidebar(){
  _sidebar.classList.add('open');
  _overlay.classList.add('open');
  _hamburger.setAttribute('aria-expanded', 'true');
  _overlay.setAttribute('aria-hidden', 'false');
}
function closeSidebar(){
  _sidebar.classList.remove('open');
  _overlay.classList.remove('open');
  _hamburger.setAttribute('aria-expanded', 'false');
  _overlay.setAttribute('aria-hidden', 'true');
}
_hamburger.addEventListener('click', () => {
  _sidebar.classList.contains('open') ? closeSidebar() : openSidebar();
});
_overlay.addEventListener('click', closeSidebar);
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeSidebar(); });

// =================================================================
// TAB 1: TIME SERIES
// =================================================================
const tsState = { symbols: [], item: null, mode: 'abs', range: 'all', type: '', industry: '' };
let tsChart = null;

function tsRebuildSymbolList(){
  const symSel = document.getElementById('ts-symbols');
  const visible = DATA.symbols.filter(s => symbolMatchesLabels(s, tsState.type, tsState.industry));
  const stillSelected = new Set(tsState.symbols.filter(s => visible.includes(s)));
  symSel.innerHTML = '';
  for (const s of visible){
    const o = document.createElement('option');
    o.value = s;
    const l = labelOf(s);
    o.textContent = l.industry ? `${s}  ·  ${l.industry}` : s;
    if (stillSelected.has(s)) o.selected = true;
    symSel.appendChild(o);
  }
  tsState.symbols = [...stillSelected];
}

function tsInit(){
  const symSel = document.getElementById('ts-symbols');
  const itemSel = document.getElementById('ts-item');
  const typeSel = document.getElementById('ts-type');
  const indSel = document.getElementById('ts-industry');
  populateLabelSelect(typeSel, TYPES);
  populateLabelSelect(indSel, INDUSTRIES);
  populateItemSelect(itemSel);
  tsRebuildSymbolList();

  // Touch devices have no Cmd/Ctrl — phrase the hint accordingly.
  const touch = ('ontouchstart' in window) || (navigator.maxTouchPoints || 0) > 0;
  const hint = document.getElementById('ts-symbols-hint');
  if (hint && touch) hint.textContent = 'Tap to add or remove symbols';

  // Defaults: first symbol + 'Total Revenue' if it exists
  const defaultItem = DATA.items.find(i => i.name === 'Total Revenue') || DATA.items[0];
  itemSel.value = defaultItem.path;
  // pre-select up to 3 of the largest stocks
  const ranked = [...DATA.symbols].sort((a,b) => {
    const va = getSeries(a, defaultItem.path);
    const vb = getSeries(b, defaultItem.path);
    const latestA = Math.max(...Object.values(va).map(Math.abs), 0);
    const latestB = Math.max(...Object.values(vb).map(Math.abs), 0);
    return latestB - latestA;
  });
  const defaultSyms = ranked.slice(0, 3);
  for (const o of symSel.options) o.selected = defaultSyms.includes(o.value);
  tsState.symbols = defaultSyms;
  tsState.item = defaultItem.path;

  symSel.addEventListener('change', () => {
    tsState.symbols = [...symSel.selectedOptions].map(o => o.value);
    tsRender();
  });
  itemSel.addEventListener('change', () => { tsState.item = itemSel.value; tsRender(); });
  typeSel.addEventListener('change', () => {
    tsState.type = typeSel.value; tsRebuildSymbolList(); tsRender();
  });
  indSel.addEventListener('change', () => {
    tsState.industry = indSel.value; tsRebuildSymbolList(); tsRender();
  });
  wireSelect('ts-mode', v => { tsState.mode = v; tsRender(); });
  wireSelect('ts-range', v => { tsState.range = v; tsRender(); });
  tsRender();
}

function tsRangeFilter(periods){
  if (tsState.range === 'all') return periods;
  const n = { '1y': 4, '2y': 8, '3y': 12 }[tsState.range] || periods.length;
  return periods.slice(-n);
}

function tsTransform(series, periods){
  // series: {period: value}, returns array aligned with periods (nullable)
  const raw = periods.map(p => series[p] != null ? series[p] : null);
  if (tsState.mode === 'abs') return raw;
  if (tsState.mode === 'qoq') {
    return periods.map((p,i) => {
      const cur = series[p], prev = series[prevQ(p)];
      return cur != null && prev != null && prev !== 0 ? (cur - prev) / Math.abs(prev) * 100 : null;
    });
  }
  if (tsState.mode === 'yoy') {
    return periods.map(p => {
      const cur = series[p], prev = series[yoyQ(p)];
      return cur != null && prev != null && prev !== 0 ? (cur - prev) / Math.abs(prev) * 100 : null;
    });
  }
  if (tsState.mode === 'idx') {
    const first = raw.find(v => v != null && v !== 0);
    if (first == null) return raw;
    return raw.map(v => v == null ? null : v / first * 100);
  }
  return raw;
}

function tsRender(){
  const allPeriods = [...DATA.periods].sort(periodSort);
  const periods = tsRangeFilter(allPeriods);
  // Colour series by industry where unique; fall back to palette for repeats.
  const usedIndColors = new Set();
  const datasets = tsState.symbols.map((s,i) => {
    const ind = labelOf(s).industry;
    const indC = INDUSTRY_COLORS[ind];
    let color;
    if (indC && !usedIndColors.has(indC)){ color = indC; usedIndColors.add(indC); }
    else color = palette[i % palette.length];
    return {
      label: ind ? `${s}  ·  ${ind}` : s,
      symbol: s,
      data: tsTransform(getSeries(s, tsState.item), periods),
      borderColor: color,
      backgroundColor: color + '1A',
      borderWidth: 1.75, tension: 0.25, spanGaps: true,
      pointRadius: 2.5, pointHoverRadius: 4,
      pointBackgroundColor: color,
      pointBorderColor: '#ffffff', pointBorderWidth: 1
    };
  });

  const ctx = document.getElementById('ts-chart');
  if (tsChart) tsChart.destroy();
  const yLabel = { abs: 'k.Baht', qoq: 'QoQ %', yoy: 'YoY %', idx: 'Index (first=100)' }[tsState.mode];
  tsChart = new Chart(ctx, {
    type: 'line',
    data: { labels: periods, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: {
          position: 'top', align: 'end',
          labels: {
            color: CHART_THEME.legendColor,
            usePointStyle: true, pointStyle: 'circle',
            boxWidth: 6, boxHeight: 6, padding: 14,
            font: { family: CHART_FONT, size: 12, weight: '500' }
          }
        },
        tooltip: Object.assign({}, TOOLTIP_STYLE, {
          callbacks: {
            label: c => {
              const sym = c.dataset.symbol || c.dataset.label;
              const p = c.label;
              const fp = fiscalOf(sym, p);
              const tag = fp && fp !== p ? `  (filed as fiscal ${fp})` : '';
              if (tsState.mode === 'abs') return `${sym}: ${fmtFull(c.parsed.y)} k.Baht${tag}`;
              if (tsState.mode === 'idx') return `${sym}: ${c.parsed.y?.toFixed(1)}${tag}`;
              return `${sym}: ${c.parsed.y?.toFixed(2)}%${tag}`;
            }
          }
        })
      },
      scales: {
        x: {
          ticks: { color: CHART_THEME.tickColor, font: { family: CHART_FONT, size: 11 } },
          grid: { color: CHART_THEME.gridColor, drawTicks: false },
          border: { color: CHART_THEME.borderColor }
        },
        y: {
          ticks: { color: CHART_THEME.tickColor, font: { family: CHART_FONT, size: 11 } },
          grid: { color: CHART_THEME.gridColor, drawTicks: false },
          border: { display: false },
          title: {
            display: true, text: yLabel,
            color: CHART_THEME.axisTitleColor,
            font: { family: CHART_FONT, size: 11, weight: '500' }
          }
        }
      }
    }
  });

  // Summary table: latest value, QoQ, YoY for selected symbols
  const latest = periods[periods.length - 1];
  const itemName = ITEMS_BY_PATH[tsState.item]?.name || tsState.item;
  let html = `<div class="panel-head"><h3>Latest snapshot</h3><span class="sub">${itemName}</span></div>`;
  html += `<div class="h-scroll"><table class="data"><thead><tr>
    <th class="sym">Symbol</th>
    <th class="item">Type</th>
    <th class="item">Industry</th>
    <th>Latest period</th>
    <th>Latest value (k.Baht)</th>
    <th>QoQ %</th>
    <th>YoY %</th>
    <th>4-yr avg (k.Baht)</th>
  </tr></thead><tbody>`;
  for (const s of tsState.symbols){
    const ser = getSeries(s, tsState.item);
    const l = labelOf(s);
    // The actual latest reported period for this stock might differ
    const symPeriods = Object.keys(ser).sort(periodSort);
    if (!symPeriods.length){
      html += `<tr><td class="sym">${s}</td><td class="item">${typePill(l.type)}</td><td class="item">${industryPill(l.industry)}</td><td colspan="5" class="empty-state">no data</td></tr>`;
      continue;
    }
    const lp = symPeriods[symPeriods.length-1];
    const lv = ser[lp];
    const pv = ser[prevQ(lp)];
    const yv = ser[yoyQ(lp)];
    const qoq = pv != null && pv !== 0 ? (lv - pv) / Math.abs(pv) * 100 : null;
    const yoy = yv != null && yv !== 0 ? (lv - yv) / Math.abs(yv) * 100 : null;
    const vals = symPeriods.map(p => ser[p]).filter(v => v != null);
    const avg = vals.length ? vals.reduce((a,b) => a+b, 0) / vals.length : null;
    const fp = fiscalOf(s, lp);
    const lpDisplay = fp && fp !== lp ? `${lp}<span class="fp-note">(fiscal ${fp})</span>` : lp;
    html += `<tr>
      <td class="sym">${s}</td>
      <td class="item">${typePill(l.type)}</td>
      <td class="item">${industryPill(l.industry)}</td>
      <td>${lpDisplay}</td>
      <td>${fmtFull(lv)}</td>
      <td class="${qoq != null && qoq >= 0 ? 'pos' : 'neg'}">${pct(qoq != null ? qoq/100 : null)}</td>
      <td class="${yoy != null && yoy >= 0 ? 'pos' : 'neg'}">${pct(yoy != null ? yoy/100 : null)}</td>
      <td>${fmtFull(avg)}</td>
    </tr>`;
  }
  html += `</tbody></table></div><div class="footnote">Item path: <b>${tsState.item}</b></div>`;
  document.getElementById('ts-summary').innerHTML = html;
}

// =================================================================
// TAB 2: PEER COMPARISON
// =================================================================
const peerState = { item: null, period: null, sort: 'desc', filter: '', compare: 'abs', type: '', industry: '' };
let peerChart = null;

function peerInit(){
  const itemSel = document.getElementById('peer-item');
  const periodSel = document.getElementById('peer-period');
  const typeSel = document.getElementById('peer-type');
  const indSel = document.getElementById('peer-industry');
  populateItemSelect(itemSel);
  populatePeriodSelect(periodSel, false);
  populateLabelSelect(typeSel, TYPES);
  populateLabelSelect(indSel, INDUSTRIES);

  const defaultItem = DATA.items.find(i => i.name === 'Total Revenue') || DATA.items[0];
  itemSel.value = defaultItem.path;
  peerState.item = defaultItem.path;
  // latest period that has data
  const latest = [...DATA.periods].sort(periodSort).reverse()[0];
  periodSel.value = latest;
  peerState.period = latest;

  itemSel.addEventListener('change', () => { peerState.item = itemSel.value; peerRender(); });
  periodSel.addEventListener('change', () => { peerState.period = periodSel.value; peerRender(); });
  typeSel.addEventListener('change', () => { peerState.type = typeSel.value; peerRender(); });
  indSel.addEventListener('change', () => { peerState.industry = indSel.value; peerRender(); });
  wireSelect('peer-sort', v => { peerState.sort = v; peerRender(); });
  wireSelect('peer-compare', v => { peerState.compare = v; peerRender(); });
  document.getElementById('peer-filter').addEventListener('input', e => {
    peerState.filter = e.target.value.toLowerCase(); peerRender();
  });
  peerRender();
}

function peerComputeRow(sym){
  const ser = getSeries(sym, peerState.item);
  const v = ser[peerState.period];
  let display = null, compareValue = null, comparePeriod = null;
  if (peerState.compare === 'abs') {
    display = v;
  } else if (peerState.compare === 'qoq') {
    comparePeriod = prevQ(peerState.period);
    compareValue = ser[comparePeriod];
    display = (v != null && compareValue != null && compareValue !== 0) ? (v - compareValue) / Math.abs(compareValue) * 100 : null;
  } else if (peerState.compare === 'yoy') {
    comparePeriod = yoyQ(peerState.period);
    compareValue = ser[comparePeriod];
    display = (v != null && compareValue != null && compareValue !== 0) ? (v - compareValue) / Math.abs(compareValue) * 100 : null;
  } else if (peerState.compare === 'rev') {
    // % of Total Revenue for that symbol/period
    const trPath = (DATA.items.find(i => i.name === 'Total Revenue') || {}).path;
    compareValue = trPath ? (DATA.values[sym]?.[trPath]?.[peerState.period]) : null;
    comparePeriod = peerState.period;
    display = (v != null && compareValue != null && compareValue !== 0) ? v / compareValue * 100 : null;
  }
  return { sym, value: display, raw: v, compareValue, comparePeriod };
}

function peerRender(){
  let rows = DATA.symbols
    .filter(s => !peerState.filter || s.toLowerCase().includes(peerState.filter))
    .filter(s => symbolMatchesLabels(s, peerState.type, peerState.industry))
    .map(peerComputeRow)
    .filter(r => r.value != null);

  if (peerState.sort === 'desc') rows.sort((a,b) => b.value - a.value);
  else if (peerState.sort === 'asc') rows.sort((a,b) => a.value - b.value);
  else rows.sort((a,b) => a.sym.localeCompare(b.sym));

  const ctx = document.getElementById('peer-chart');
  if (peerChart) peerChart.destroy();
  // Size the chart's container to the row count so each bar gets enough
  // vertical room for its label. With ~56 symbols the fixed 440-620px
  // container squashes ticks into illegible overlap.
  const wrap = ctx.parentElement;
  const perRow = 22;
  const minHeight = window.matchMedia('(max-width: 640px)').matches ? 280 : 420;
  wrap.style.height = Math.max(minHeight, rows.length * perRow + 60) + 'px';
  const isPct = peerState.compare !== 'abs';
  // Always colour bars by industry — sign of % is conveyed by the bar
  // direction (negative bars extend left of zero) and by the green/red
  // text in the table below.
  const colors = rows.map(r => industryColor(labelOf(r.sym).industry));
  const hovers = colors.map(c => c);
  peerChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: rows.map(r => r.sym),
      datasets: [{
        label: ITEMS_BY_PATH[peerState.item].name,
        data: rows.map(r => r.value),
        backgroundColor: colors,
        hoverBackgroundColor: hovers,
        borderWidth: 0,
        borderRadius: 3,
        barThickness: 'flex',
        maxBarThickness: 18,
        categoryPercentage: 0.85,
        barPercentage: 0.9,
      }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: Object.assign({}, TOOLTIP_STYLE, {
          callbacks: {
            label: c => {
              const sym = c.label;
              const fp = fiscalOf(sym, peerState.period);
              const tag = fp && fp !== peerState.period ? `  (filed as fiscal ${fp})` : '';
              return (isPct ? `${c.parsed.x.toFixed(2)}%` : `${fmtFull(c.parsed.x)} k.Baht`) + tag;
            }
          }
        })
      },
      scales: {
        x: {
          ticks: { color: CHART_THEME.tickColor, font: { family: CHART_FONT, size: 11 } },
          grid: { color: CHART_THEME.gridColor, drawTicks: false },
          border: { display: false }
        },
        y: {
          ticks: {
            color: CHART_THEME.tickColor, autoSkip: false,
            font: { family: CHART_FONT, size: 11, weight: '500' }
          },
          grid: { color: 'transparent', drawTicks: false },
          border: { color: CHART_THEME.borderColor }
        }
      }
    }
  });

  // Table
  const itemName = ITEMS_BY_PATH[peerState.item]?.name;
  const headerVal = {
    abs: `${itemName} (k.Baht)`,
    qoq: `QoQ %`, yoy: `YoY %`, rev: `% of Total Revenue`
  }[peerState.compare];
  // Build header. When comparing, expose both the current Q and the comparison
  // Q raw values so the % delta is easy to interpret at a glance.
  const cmp = peerState.compare;
  const curQ = peerState.period;
  let html = `<thead><tr>
    <th class="sym">Rank</th><th class="sym">Symbol</th>
    <th class="item">Type</th>
    <th class="item">Industry</th>
    <th>${headerVal}</th>`;
  if (cmp === 'abs') {
    html += `<th>Raw value (k.Baht)</th>`;
  } else if (cmp === 'qoq') {
    html += `<th>${curQ} (k.Baht)</th><th>${prevQ(curQ)} (k.Baht)</th>`;
  } else if (cmp === 'yoy') {
    html += `<th>${curQ} (k.Baht)</th><th>${yoyQ(curQ)} (k.Baht)</th>`;
  } else if (cmp === 'rev') {
    html += `<th>${itemName} (k.Baht)</th><th>Total Revenue (k.Baht)</th>`;
  }
  html += `</tr></thead><tbody>`;
  rows.forEach((r,i) => {
    const fmtVal = isPct ? pct(r.value/100) : fmtFull(r.value);
    const cls = isPct ? (r.value >= 0 ? 'pos' : 'neg') : '';
    const l = labelOf(r.sym);
    html += `<tr>
      <td class="sym">${i+1}</td>
      <td class="sym">${r.sym}</td>
      <td class="item">${typePill(l.type)}</td>
      <td class="item">${industryPill(l.industry)}</td>
      <td class="${cls}">${fmtVal}</td>`;
    if (cmp === 'abs') {
      html += `<td>${fmtFull(r.raw)}</td>`;
    } else {
      html += `<td>${fmtFull(r.raw)}</td><td>${fmtFull(r.compareValue)}</td>`;
    }
    html += `</tr>`;
  });
  html += '</tbody>';
  document.getElementById('peer-table').innerHTML = html;
}

// =================================================================
// TAB 3: HEATMAP
// =================================================================
// sortCol overrides the preset .sort when set: 'symbol' or a period like '2025Q4'.
const heatState = { item: null, mode: 'qoq', filter: '', sort: 'alpha', type: '', industry: '',
                    sortCol: null, sortAsc: true };

function heatInit(){
  const itemSel = document.getElementById('heat-item');
  const typeSel = document.getElementById('heat-type');
  const indSel = document.getElementById('heat-industry');
  populateItemSelect(itemSel);
  populateLabelSelect(typeSel, TYPES);
  populateLabelSelect(indSel, INDUSTRIES);
  const defaultItem = DATA.items.find(i => i.name === 'Total Revenue') || DATA.items[0];
  itemSel.value = defaultItem.path;
  heatState.item = defaultItem.path;
  itemSel.addEventListener('change', () => { heatState.item = itemSel.value; heatRender(); });
  typeSel.addEventListener('change', () => { heatState.type = typeSel.value; heatRender(); });
  indSel.addEventListener('change', () => { heatState.industry = indSel.value; heatRender(); });
  wireSelect('heat-mode', v => { heatState.mode = v; heatRender(); });
  wireSelect('heat-sort', v => { heatState.sort = v; heatState.sortCol = null; heatRender(); });
  document.getElementById('heat-filter').addEventListener('input', e => {
    heatState.filter = e.target.value.toLowerCase(); heatRender();
  });
  heatRender();
}

function heatCellValue(sym, period){
  const ser = getSeries(sym, heatState.item);
  const v = ser[period];
  if (heatState.mode === 'abs') return v;
  if (heatState.mode === 'qoq') {
    const pv = ser[prevQ(period)];
    return (v != null && pv != null && pv !== 0) ? (v - pv) / Math.abs(pv) * 100 : null;
  }
  if (heatState.mode === 'yoy') {
    const yv = ser[yoyQ(period)];
    return (v != null && yv != null && yv !== 0) ? (v - yv) / Math.abs(yv) * 100 : null;
  }
}

function heatColor(v, maxAbs){
  if (v == null || isNaN(v) || maxAbs === 0) return 'transparent';
  const t = Math.max(-1, Math.min(1, v / maxAbs));
  if (t >= 0) {
    // green — light-mode friendly: subtle baseline + saturating ramp
    const a = (0.10 + 0.50 * t).toFixed(3);
    return `rgba(34,197,94,${a})`;
  } else {
    const a = (0.10 + 0.50 * -t).toFixed(3);
    return `rgba(239,68,68,${a})`;
  }
}

function heatRender(){
  const periods = [...DATA.periods].sort(periodSort);
  const syms = DATA.symbols
    .filter(s => !heatState.filter || s.toLowerCase().includes(heatState.filter))
    .filter(s => symbolMatchesLabels(s, heatState.type, heatState.industry));

  // Sort: an explicit column click overrides the preset dropdown
  if (heatState.sortCol === 'symbol'){
    syms.sort((a,b) => a.localeCompare(b));
    if (!heatState.sortAsc) syms.reverse();
  } else if (heatState.sortCol){
    const col = heatState.sortCol;
    syms.sort((a,b) => {
      const va = heatCellValue(a, col), vb = heatCellValue(b, col);
      if (va == null && vb == null) return 0;
      if (va == null) return 1;   // nulls always sink to bottom
      if (vb == null) return -1;
      return heatState.sortAsc ? va - vb : vb - va;
    });
  } else if (heatState.sort === 'last') {
    const latest = periods[periods.length-1];
    syms.sort((a,b) => (heatCellValue(b,latest)||-Infinity) - (heatCellValue(a,latest)||-Infinity));
  } else if (heatState.sort === 'latestqoq') {
    const latest = periods[periods.length-1];
    const tmp = heatState.mode; heatState.mode = 'qoq';
    syms.sort((a,b) => (heatCellValue(b,latest)||-Infinity) - (heatCellValue(a,latest)||-Infinity));
    heatState.mode = tmp;
  } else {
    syms.sort((a,b) => a.localeCompare(b));
  }

  // Compute matrix
  const matrix = syms.map(s => periods.map(p => heatCellValue(s,p)));

  // Scale: cap at 200% for QoQ/YoY for legibility; or 95th percentile of abs for absolute
  let maxAbs = 0;
  if (heatState.mode === 'abs') {
    const flat = matrix.flat().filter(v => v != null).map(Math.abs).sort((a,b) => a-b);
    maxAbs = flat.length ? flat[Math.floor(flat.length * 0.95)] : 1;
  } else {
    maxAbs = 50; // 50% change = saturated colour
  }
  if (maxAbs === 0) maxAbs = 1;

  const symSortCls = heatState.sortCol === 'symbol' ? (heatState.sortAsc ? 'sort-asc' : 'sort-desc') : '';
  let html = `<thead><tr><th class="sym ${symSortCls}" data-col="symbol">Symbol</th>`;
  for (const p of periods) {
    const sortCls = heatState.sortCol === p ? (heatState.sortAsc ? 'sort-asc' : 'sort-desc') : '';
    html += `<th class="${sortCls}" data-col="${p}">${periodLabel(p)}</th>`;
  }
  html += `</tr></thead><tbody>`;
  syms.forEach((s,i) => {
    const l = labelOf(s);
    const dot = `<span class="ind-dot" style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${industryColor(l.industry)};margin-right:7px;vertical-align:middle"></span>`;
    const titleAttr = l.industry ? ` title="${l.type ? l.type + ' · ' : ''}${l.industry}"` : '';
    html += `<tr><td class="sym"${titleAttr}>${dot}${s}</td>`;
    matrix[i].forEach(v => {
      const bg = heatColor(v, maxAbs);
      const txt = v == null ? '' : (heatState.mode === 'abs' ? fmt(v) : v.toFixed(1) + '%');
      const cls = v == null ? 'empty' : '';
      html += `<td class="${cls}" style="background:${bg}">${txt}</td>`;
    });
    html += `</tr>`;
  });
  html += `</tbody>`;
  const table = document.getElementById('heat-table');
  table.innerHTML = html;
  // Wire click-to-sort on every column header. Default direction: ascending for
  // the symbol column (A→Z), descending for value columns (largest first).
  table.querySelectorAll('thead th').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      if (heatState.sortCol === col) heatState.sortAsc = !heatState.sortAsc;
      else { heatState.sortCol = col; heatState.sortAsc = (col === 'symbol'); }
      heatRender();
    });
  });
}

// =================================================================
// TAB 4: RAW DATA
// =================================================================
let rawSortKey = 'symbol', rawSortAsc = true;
let rawAllRows = null;

function rawFlatRows(){
  if (rawAllRows) return rawAllRows;
  // Some item_paths in DATA.values may have been hidden from ITEMS_BY_PATH
  // (de-duplicated OCI variants). Recover their display name + category by
  // matching the leaf name.
  const byLeaf = {};
  for (const it of DATA.items) byLeaf[it.name] = it;
  const out = [];
  for (const sym in DATA.values) {
    const byItem = DATA.values[sym];
    const l = labelOf(sym);
    for (const ip in byItem) {
      let meta = ITEMS_BY_PATH[ip];
      if (!meta){
        // dropped duplicate — synthesize meta from the path
        const leaf = ip.split(' > ').pop();
        meta = byLeaf[leaf] || { name: leaf, level: 3, category: 'Other' };
      }
      for (const p in byItem[ip]) {
        out.push({
          symbol: sym, period: p, fiscal_period: fiscalOf(sym, p) || p,
          item: meta.name, item_path: ip, level: meta.level,
          category: meta.category || 'Other',
          type: l.type || '', industry: l.industry || '',
          value: byItem[ip][p]
        });
      }
    }
  }
  rawAllRows = out;
  return out;
}

function rawInit(){
  populatePeriodSelect(document.getElementById('raw-period'), true);
  populateLabelSelect(document.getElementById('raw-type'), TYPES);
  populateLabelSelect(document.getElementById('raw-industry'), INDUSTRIES);
  ['raw-sym','raw-item','raw-min'].forEach(id => {
    document.getElementById(id).addEventListener('input', rawRender);
  });
  document.getElementById('raw-period').addEventListener('change', rawRender);
  document.getElementById('raw-cat').addEventListener('change', rawRender);
  document.getElementById('raw-type').addEventListener('change', rawRender);
  document.getElementById('raw-industry').addEventListener('change', rawRender);
  rawRender();
}

function rawRender(){
  const symF   = document.getElementById('raw-sym').value.trim().toLowerCase();
  const itemF  = document.getElementById('raw-item').value.trim().toLowerCase();
  const perF   = document.getElementById('raw-period').value;
  const catF   = document.getElementById('raw-cat').value;
  const typeF  = document.getElementById('raw-type').value;
  const indF   = document.getElementById('raw-industry').value;
  const minF   = parseFloat(document.getElementById('raw-min').value);

  let rows = rawFlatRows();
  rows = rows.filter(r =>
    (!symF || r.symbol.toLowerCase().includes(symF)) &&
    (!itemF || r.item.toLowerCase().includes(itemF) || r.item_path.toLowerCase().includes(itemF)) &&
    (!perF || r.period === perF) &&
    (!catF || r.category === catF) &&
    (!typeF || r.type === typeF) &&
    (!indF || r.industry === indF) &&
    (isNaN(minF) || Math.abs(r.value) >= minF)
  );

  rows.sort((a,b) => {
    const A = a[rawSortKey], B = b[rawSortKey];
    let cmp;
    if (typeof A === 'number' && typeof B === 'number') cmp = A - B;
    else cmp = String(A).localeCompare(String(B));
    return rawSortAsc ? cmp : -cmp;
  });

  document.getElementById('raw-count').textContent = `${rows.length.toLocaleString()} rows`;

  const cap = 2000;
  const shown = rows.slice(0, cap);
  const cols = [
    ['symbol','Symbol','sym'],
    ['type','Type','item'],
    ['industry','Industry','item'],
    ['period','Calendar Q',''], ['fiscal_period','Fiscal Q',''],
    ['category','Category','item'], ['item','Item','item'], ['item_path','Item path','item'],
    ['value','Value (k.Baht)',''],
  ];
  let html = `<thead><tr>`;
  for (const [k,label,cls] of cols){
    const sortCls = rawSortKey === k ? (rawSortAsc ? 'sort-asc' : 'sort-desc') : '';
    html += `<th class="${cls} ${sortCls}" data-k="${k}">${label}</th>`;
  }
  html += `</tr></thead><tbody>`;
  for (const r of shown){
    const fpCls = r.fiscal_period !== r.period
      ? 'style="color:#b45309;font-weight:500"'
      : 'style="color:var(--muted)"';
    html += `<tr>
      <td class="sym">${r.symbol}</td>
      <td class="item">${typePill(r.type)}</td>
      <td class="item">${industryPill(r.industry)}</td>
      <td>${r.period}</td>
      <td ${fpCls}>${r.fiscal_period}</td>
      <td class="item"><span class="cat-pill cat-${(r.category||'').replace(/[^a-z]/gi,'')}">${r.category||''}</span></td>
      <td class="item">${r.item}</td>
      <td class="item" style="color:var(--muted)">${r.item_path}</td>
      <td>${fmtFull(r.value)}</td>
    </tr>`;
  }
  if (rows.length > cap) {
    html += `<tr><td colspan="${cols.length}" style="color:var(--muted);text-align:center;padding:10px">…showing first ${cap.toLocaleString()} of ${rows.length.toLocaleString()} rows — refine filter to see more</td></tr>`;
  }
  html += `</tbody>`;
  const table = document.getElementById('raw-table');
  table.innerHTML = html;
  table.querySelectorAll('th').forEach(th => {
    th.addEventListener('click', () => {
      const k = th.dataset.k;
      if (rawSortKey === k) rawSortAsc = !rawSortAsc;
      else { rawSortKey = k; rawSortAsc = true; }
      rawRender();
    });
  });
}

// ----- boot -----
tsInit();
peerInit();
heatInit();
rawInit();
</script>
</body>
</html>
"""


def render_html(payload: dict) -> str:
    data_json = json.dumps(payload, separators=(",", ":"), ensure_ascii=False)
    # Make sure no `</script>` sneaks in via item names
    data_json = data_json.replace("</", "<\\/")
    html = HTML_TEMPLATE
    html = html.replace("__DATA_JSON__", data_json)
    html = html.replace("__SYMBOL_COUNT__", str(len(payload["symbols"])))
    html = html.replace("__PERIOD_COUNT__", str(len(payload["periods"])))
    html = html.replace("__ROW_COUNT__", f"{payload['meta']['row_count']:,}")
    html = html.replace("__GENERATED__", payload["meta"]["generated"])
    return html


def main() -> int:
    rows = load_csv()
    labels = load_labels()
    payload = build_payload(rows, labels)
    html = render_html(payload)
    HTML_PATH.write_text(html, encoding="utf-8")
    size_kb = HTML_PATH.stat().st_size / 1024
    print(f"Wrote {HTML_PATH.relative_to(ROOT)}  ({size_kb:,.1f} KB, {payload['meta']['row_count']:,} data points)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
