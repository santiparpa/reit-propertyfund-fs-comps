"""
build_db.py
-----------
Scan all *_FS_*.xlsx in the project root and emit a single normalized CSV
(`data/database.csv`) in long format:

    symbol, period, period_end, item, item_path, level, value

Re-runnable: drop new quarterly xlsx files in the root and run this script
again to refresh the database from scratch.
"""

from __future__ import annotations

import csv
import glob
import os
import re
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is required. Install with: pip install openpyxl")


ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "raw"
DATA_DIR = ROOT / "data"
OUT_CSV = DATA_DIR / "database.csv"

# Quarter column label, e.g. " Q1 / 2026  (01/01/26 - 31/03/26) Company"
QUARTER_RE = re.compile(
    r"Q(?P<q>\d)\s*/\s*(?P<year>\d{4}).*?\((?P<start>\d{2}/\d{2}/\d{2})\s*-\s*(?P<end>\d{2}/\d{2}/\d{2})\)",
    re.IGNORECASE,
)

# Rows past the data block. We stop reading items when we hit any of these.
TERMINATORS = {
    "Financial Statement (Full Version):",
    "Remark:",
}


def parse_quarter_label(label: str) -> tuple[str, str, str] | None:
    """
    Return (calendar_period, period_end_iso, fiscal_period) or None.

    `calendar_period` is derived from the period_end month (Jan-Mar→Q1, etc.) so
    REITs with non-calendar fiscal years are aligned to calendar quarters.
    `fiscal_period` is the company's own Q-labelling (kept for transparency).
    """
    if not isinstance(label, str):
        return None
    m = QUARTER_RE.search(label)
    if not m:
        return None
    fiscal_period = f"{m['year']}Q{m['q']}"
    dd, mm, yy = m["end"].split("/")
    period_end = f"20{yy}-{mm}-{dd}"
    cal_year = 2000 + int(yy)
    cal_q = (int(mm) - 1) // 3 + 1
    calendar_period = f"{cal_year}Q{cal_q}"
    return calendar_period, period_end, fiscal_period


def detect_level(raw: str) -> int:
    """
    Item indent → hierarchy level:
      1 space  → level 1 (section header, usually empty values)
      4 spaces → level 2 (main line item)
      6 spaces → level 3 (sub-component)
    Anything else → 0 (not a data row).
    """
    leading = len(raw) - len(raw.lstrip(" "))
    if leading == 1:
        return 1
    if leading == 4:
        return 2
    if leading == 6:
        return 3
    return 0


def to_number(v) -> float | None:
    """Numbers pass through; the ' - ' sentinel and blanks become None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s in {"-", "–", "—"}:
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def parse_file(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Find symbol in B3 (row index 2, col index 1)
    symbol = str(rows[2][1]).strip() if rows[2][1] else path.stem.split("_")[0]

    # Locate the quarter-header row: first row whose col B matches the quarter pattern.
    header_idx = None
    for i, r in enumerate(rows):
        if len(r) > 1 and parse_quarter_label(r[1] or ""):
            header_idx = i
            break
    if header_idx is None:
        print(f"WARN  {path.name}: no quarter header row found, skipping")
        return []

    # Build column → (calendar_period, period_end, fiscal_period) map
    col_to_period: dict[int, tuple[str, str, str]] = {}
    for col_idx, cell in enumerate(rows[header_idx]):
        if col_idx == 0:
            continue
        parsed = parse_quarter_label(cell or "")
        if parsed:
            col_to_period[col_idx] = parsed

    if not col_to_period:
        print(f"WARN  {path.name}: header row had no parseable quarters")
        return []

    out: list[dict] = []
    # Walk through item rows, tracking parent section / parent item
    parent_by_level: dict[int, str] = {}
    for r in rows[header_idx + 1 :]:
        raw_item = r[0]
        if raw_item is None:
            continue
        if not isinstance(raw_item, str):
            continue
        if raw_item.strip() in TERMINATORS:
            break
        # Skip the disclaimer lines
        if raw_item.startswith("*") or raw_item.startswith("Information"):
            continue
        if raw_item.startswith("Restatement"):
            continue

        level = detect_level(raw_item)
        if level == 0:
            continue

        name = raw_item.strip()
        parent_by_level[level] = name
        # Invalidate deeper levels when we ascend
        for deeper in [l for l in parent_by_level if l > level]:
            del parent_by_level[deeper]

        # Build item_path from levels 1..current
        path_parts = [parent_by_level[l] for l in sorted(parent_by_level) if l <= level]
        item_path = " > ".join(path_parts)

        # Section headers (level 1) have no values; record them only if any column is non-empty
        # In practice level-1 rows are always blank, so we just skip them.
        if level == 1:
            continue

        for col_idx, (period, period_end, fiscal_period) in col_to_period.items():
            raw_val = r[col_idx] if col_idx < len(r) else None
            val = to_number(raw_val)
            if val is None:
                continue  # don't bloat the CSV with nulls
            out.append(
                {
                    "symbol": symbol,
                    "period": period,
                    "period_end": period_end,
                    "fiscal_period": fiscal_period,
                    "item": name,
                    "item_path": item_path,
                    "level": level,
                    "value": val,
                }
            )
    return out


def main() -> int:
    files = sorted(glob.glob(str(RAW_DIR / "*_FS_*.xlsx")))
    if not files:
        sys.exit(
            f"No *_FS_*.xlsx files found in {RAW_DIR}\n"
            f"Drop new quarterly SETSmart exports into raw/ and re-run."
        )

    DATA_DIR.mkdir(exist_ok=True)

    all_rows: list[dict] = []
    per_symbol_count: dict[str, int] = {}
    for f in files:
        rows = parse_file(Path(f))
        all_rows.extend(rows)
        sym = Path(f).stem.split("_")[0]
        per_symbol_count[sym] = len(rows)

    # Stable ordering: symbol → period (desc) → item_path
    all_rows.sort(key=lambda r: (r["symbol"], r["period"], r["item_path"]), reverse=False)

    with OUT_CSV.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(
            fh,
            fieldnames=[
                "symbol", "period", "period_end", "fiscal_period",
                "item", "item_path", "level", "value",
            ],
        )
        w.writeheader()
        w.writerows(all_rows)

    print(f"Wrote {OUT_CSV.relative_to(ROOT)}  ({len(all_rows):,} rows from {len(files)} files)")

    # Brief sanity report
    empty = [s for s, n in per_symbol_count.items() if n == 0]
    if empty:
        print(f"WARN  {len(empty)} symbol(s) produced 0 rows: {', '.join(empty)}")
    sizes = sorted(per_symbol_count.values())
    if sizes:
        print(f"Rows per symbol — min={sizes[0]}, median={sizes[len(sizes)//2]}, max={sizes[-1]}")

    # Fiscal-year offset report: which symbols have non-calendar fiscal years?
    #   offset_quarters = fiscal_quarter - calendar_quarter for the same period_end
    fiscal_offsets: dict[str, set] = {}
    for r in all_rows:
        fy, fq = r["fiscal_period"].split("Q")
        cy, cq = r["period"].split("Q")
        # how many calendar quarters does the fiscal label lead the calendar quarter by?
        offset = (int(fy) * 4 + int(fq)) - (int(cy) * 4 + int(cq))
        fiscal_offsets.setdefault(r["symbol"], set()).add(offset)

    non_calendar = {s: list(offsets)[0] for s, offsets in fiscal_offsets.items()
                    if offsets != {0}}
    if non_calendar:
        print("\nNon-calendar fiscal years (data realigned to calendar quarters):")
        for s, off in sorted(non_calendar.items()):
            month = {1: "Oct", 2: "Jul", 3: "Apr"}.get(off, f"+{off}Q")
            print(f"  {s:8s}  fiscal year starts in {month}  (fiscal Q is {off} quarters ahead of calendar)")
    else:
        print("\nAll symbols use calendar-aligned fiscal years.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
