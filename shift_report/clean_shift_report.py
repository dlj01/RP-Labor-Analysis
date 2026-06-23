#!/usr/bin/env python3
"""
clean_shift_report.py
---------------------
Takes a Sling/Toast shift-detail export (report-export.csv) and:
 
  1. Drops all rows belonging to salaried employees (read from
     SALARIED_EMPLOYEES.txt, with name-variant matching).
  2. Drops the per-employee summary/total rows that sit above each
     employee's block -- they go stale once shifts are combined. With a
     SOURCE column these are the blank-SOURCE rows; without one, they are
     identified as the rows that have no DATE.
  3. Combines rows that share the same EMPLOYEE + DATE + POSITIONS into a
     single row, regardless of source. Duplicates arise because each shift
     is reported twice -- once from Toast and once from Sling -- so the two
     rows are merged into one. Scheduled and actual durations are summed and
     the clock window is spanned (earliest in -> latest out). SOURCE, when
     present, is used to drop the summary rows in step 1 but is not kept in
     the output.
  4. Drops the LEGAL NAME, SOURCE, and NOTES columns (whichever are present)
     and appends a VARIANCE FLAG column:
       IN PROGRESS -> no completed actual shift yet (scheduled but not worked,
                      or clocked in with no clock-out)
       OVER        -> worked 2.00+ hours over schedule
       UNDER       -> worked 2.00+ hours under schedule
       (blank)     -> worked within 2 hours of schedule
 
This script tolerates an input CSV that has no SOURCE column.
 
Note: DIFFERENCE is recomputed on the merged row as actual - scheduled so it
stays consistent with the combined durations (the per-row differences are
not simply added, since a Sling row contributes scheduled hours with no
matching actual).
 
Inputs
------
Drop this script in the same folder as build_payroll_with_scheduled.py and the
data files, then run it with no arguments. It auto-discovers, searching the
script's own folder first and then its parent folder:
 
  report-export*.csv      the Sling/Toast shift-detail export (input)
  SALARIED_EMPLOYEES.txt   salaried-employee list (also accepts the
                           "SALARIED EMPLOYEES.txt" spelling)
 
Any path can still be overridden explicitly:
 
  python3 clean_shift_report.py [INPUT.csv] [-s SALARIED.txt] [-o OUTPUT.csv]
 
Output defaults to report-export_cleaned.csv next to the input file. The
auto-discovery deliberately ignores any existing *_cleaned.csv so a re-run
never picks up its own previous output as the input.
"""
 
import argparse
import re
import sys
from pathlib import Path
 
import pandas as pd
 
# --- column names as they appear in the export (newlines preserved) ---
EMP   = "EMPLOYEE"
LEGAL = "LEGAL NAME"
DATE  = "DATE"
POS   = "POSITIONS"
SRC   = "SOURCE"
NOTES = "NOTES"
SCH_START = "SCH.\nSHIFT START"
SCH_END   = "SCH.\nSHIFT END"
SCH_DUR   = "SCH. SHIFT\nDURATION"
CLK_IN    = "CLOCK IN\nTIME"
CLK_OUT   = "CLOCK OUT\nTIME"
SHIFT_DUR = "SHIFT\nDURATION"
DIFF      = "DIFFERENCE"
 
KEY = [EMP, DATE, POS]
 
 
# ---------------------------------------------------------------------------
# Name normalization + salaried matching
# ---------------------------------------------------------------------------
def normalize_name(name: str) -> frozenset:
    """Normalize a name to a set of significant tokens (order-independent).
 
    Handles 'Last, First A.' ordering, strips punctuation and single-letter
    middle initials, lowercases. Returns a frozenset of tokens so name order
    and middle names don't matter.
    """
    name = (name or "").strip()
    if "," in name:                         # 'Diaz Martinez, Elmer A.'
        last, first = name.split(",", 1)
        name = f"{first} {last}"
    name = re.sub(r"[.,]", " ", name).lower()
    return frozenset(t for t in name.split() if len(t) > 1)  # drop initials
 
 
def load_salaried(path: str) -> list:
    """Read SALARIED_EMPLOYEES.txt -> list of normalized token sets.
 
    Expects lines like 'Name - Role'. Ignores blanks, headers, and dividers.
    """
    sets = []
    with open(path, encoding="utf-8-sig") as fh:
        for line in fh:
            if " - " in line:
                sets.append(normalize_name(line.split(" - ", 1)[0]))
    return [s for s in sets if s]
 
 
def is_salaried(employee: str, salaried_sets: list) -> bool:
    """True if the employee matches any salaried name.
 
    Match when one token set is a subset of the other (either direction),
    which catches 'Jose Roberto Carlos Nolasco Diaz' vs 'Roberto Nolasco
    Diaz' and 'Diaz Martinez, Elmer A.' vs 'Elmer Andres Diaz Martinez',
    without sweeping up unrelated people who merely share a last name
    (e.g. 'Mabel Diaz' will NOT match 'Jose Misael Diaz').
    """
    e = normalize_name(employee)
    if not e:
        return False
    return any(s <= e or e <= s for s in salaried_sets)
 
 
# ---------------------------------------------------------------------------
# Time + number helpers for combining
# ---------------------------------------------------------------------------
def to_minutes(t: str):
    """Parse 'H:MM AM/PM' -> minutes since midnight. Returns None if blank/'-'."""
    t = (t or "").strip()
    if t in ("", "-"):
        return None
    m = re.match(r"(\d{1,2}):(\d{2})\s*([AP]M)", t, re.IGNORECASE)
    if not m:
        return None
    hh, mm, ap = int(m.group(1)), int(m.group(2)), m.group(3).upper()
    if ap == "PM" and hh != 12:
        hh += 12
    if ap == "AM" and hh == 12:
        hh = 0
    return hh * 60 + mm
 
 
def to_num(x):
    """Parse a duration string to float. Returns None if blank/'-'."""
    x = (x or "").strip()
    if x in ("", "-"):
        return None
    try:
        return float(x)
    except ValueError:
        return None
 
 
def earliest_time(values):
    """Return the time string that is earliest in the day (None-safe)."""
    cand = [(to_minutes(v), v) for v in values if to_minutes(v) is not None]
    return min(cand, key=lambda p: p[0])[1] if cand else ""
 
 
def latest_time(values):
    """Return the latest clock time; treats early-AM (<6 AM) as past midnight
    so a 1:10 AM late-night closeout ranks after a PM clock-out."""
    def rank(v):
        m = to_minutes(v)
        return m + 1440 if (m is not None and m < 360) else m
    cand = [(rank(v), v) for v in values if to_minutes(v) is not None]
    return max(cand, key=lambda p: p[0])[1] if cand else ""
 
 
def sum_num(values, decimals=2):
    """Sum the numeric values; '' if none are numeric."""
    nums = [to_num(v) for v in values]
    nums = [n for n in nums if n is not None]
    return f"{sum(nums):.{decimals}f}" if nums else ""
 
 
def first_nonblank(values):
    for v in values:
        if str(v).strip():
            return v
    return ""
 
 
def join_notes(values):
    seen, out = set(), []
    for v in values:
        v = str(v).strip()
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return " | ".join(out)
 
 
# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def clean(input_csv: str, salaried_txt: str, output_csv: str) -> None:
    df = pd.read_csv(input_csv, dtype=str).fillna("")
    original_cols = list(df.columns)
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip()
 
    # 1. keep only real shift-detail rows, dropping the per-employee
    #    summary/total rows that sit above each employee's block. When the
    #    export has a SOURCE column those summary rows are the blank-SOURCE
    #    ones (everything else is Toast/Sling). Without a SOURCE column, fall
    #    back to dropping rows that have no DATE -- the summary rows leave
    #    DATE (and POSITIONS) blank, while every real shift has a date.
    before = len(df)
    if SRC in df.columns:
        df = df[df[SRC].isin(["Toast", "Sling"])].copy()
    elif DATE in df.columns:
        df = df[df[DATE].str.strip() != ""].copy()
    dropped_summary = before - len(df)
 
    # 2. drop salaried employees
    salaried_sets = load_salaried(salaried_txt)
    sal_mask = df[EMP].apply(lambda e: is_salaried(e, salaried_sets))
    dropped_names = sorted(df.loc[sal_mask, EMP].unique())
    df = df[~sal_mask].copy()
 
    # 3. combine rows sharing EMPLOYEE + DATE + POSITIONS (across sources)
    agg = {
        SCH_START: earliest_time,
        SCH_END:   latest_time,
        SCH_DUR:   sum_num,
        CLK_IN:    earliest_time,
        CLK_OUT:   latest_time,
        SHIFT_DUR: sum_num,
    }
    combined = (
        df.groupby(KEY, sort=False, as_index=False)
          .agg({col: fn for col, fn in agg.items() if col in df.columns})
    )
 
    # recompute DIFFERENCE on the merged row = actual - scheduled, so it
    # stays consistent with the summed durations
    def recompute_diff(row):
        a, s = to_num(row[SHIFT_DUR]), to_num(row[SCH_DUR])
        if a is None and s is None:
            return ""
        return f"{(a or 0) - (s or 0):.2f}"
    combined[DIFF] = combined.apply(recompute_diff, axis=1)
 
    # flag column:
    #   IN PROGRESS -> no completed actual shift yet (scheduled but not worked,
    #                  or clocked in with no clock-out)
    #   OVER        -> worked 2.00+ hours over schedule
    #   UNDER       -> worked 2.00+ hours under schedule
    #   blank       -> worked within 2 hours of schedule
    FLAG = "VARIANCE FLAG"
    def variance_flag(row):
        if to_num(row[SHIFT_DUR]) is None:      # no completed actual punch
            return "IN PROGRESS"
        n = to_num(row[DIFF])
        if n is None:
            return ""
        if n >= 2.0:
            return "OVER"
        if n <= -2.0:
            return "UNDER"
        return ""
    combined[FLAG] = combined.apply(variance_flag, axis=1)
 
    # drop LEGAL NAME, SOURCE, and NOTES; keep original column order, append flag
    drop_cols = {LEGAL, SRC, NOTES}
    out_cols = [c for c in original_cols if c not in drop_cols] + [FLAG]
    combined = combined.reindex(columns=out_cols).fillna("")
    combined.to_csv(output_csv, index=False)
 
    # --- run report (diagnostic) ---
    print(f"Input rows:                 {before}")
    print(f"Summary/total rows dropped: {dropped_summary}")
    print(f"Salaried employees dropped: {len(dropped_names)}")
    for n in dropped_names:
        print(f"    - {n}")
    print(f"Rows after combining:       {len(combined)}")
    print(f"Wrote: {output_csv}")
 
 
def main():
    script_dir = Path(__file__).resolve().parent
    search_dirs = [script_dir, script_dir.parent]   # own folder, then parent
 
    def find_input():
        # prefer an exact 'report-export.csv', else newest 'report-export*.csv',
        # skipping any cleaned output so a re-run doesn't eat its own result
        for folder in search_dirs:
            exact = folder / "report-export.csv"
            if exact.exists():
                return exact
        for folder in search_dirs:
            matches = [m for m in sorted(folder.glob("report-export*.csv"))
                       if "cleaned" not in m.name.lower()]
            if matches:
                return matches[-1]
        return None
 
    def find_salaried():
        for folder in search_dirs:
            for name in ("SALARIED_EMPLOYEES.txt", "SALARIED EMPLOYEES.txt"):
                p = folder / name
                if p.exists():
                    return p
        return None
 
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("input_csv", nargs="?", default=None,
                    help="Shift-detail export CSV (default: auto-discovered)")
    ap.add_argument("-s", "--salaried", default=None,
                    help="SALARIED_EMPLOYEES.txt (default: auto-discovered)")
    ap.add_argument("-o", "--output", default=None,
                    help="Output CSV (default: report-export_cleaned.csv next to input)")
    args = ap.parse_args()
 
    input_csv = Path(args.input_csv) if args.input_csv else find_input()
    salaried = Path(args.salaried) if args.salaried else find_salaried()
 
    missing = []
    if input_csv is None or not input_csv.exists():
        missing.append("  input CSV (report-export*.csv)")
    if salaried is None or not salaried.exists():
        missing.append("  SALARIED_EMPLOYEES.txt")
    if missing:
        sys.stderr.write("ERROR: required input file(s) not found:\n"
                         + "\n".join(missing) + "\n")
        sys.stderr.write("Searched these folders, in order:\n"
                         + "\n".join(f"  {d}" for d in search_dirs) + "\n")
        sys.exit(1)
 
    output = (Path(args.output) if args.output
              else input_csv.with_name("report-export_cleaned.csv"))
 
    try:
        clean(str(input_csv), str(salaried), str(output))
    except FileNotFoundError as e:
        sys.exit(f"File not found: {e.filename}")
 
 
if __name__ == "__main__":
    main()
 