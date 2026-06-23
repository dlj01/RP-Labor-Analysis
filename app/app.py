#!/usr/bin/env python3
"""
l10_labor_cogs_app.py — Weekly Labor & COGS summarizer for the Rosslyn L10.

One command turns a week's raw exports into the Labor & COGS section: Sales,
the full Labor breakdown (AN/RP/Salaried, FOH/BOH, RP BOH Assembly vs Prep),
Scheduled-vs-Actual variance, COGS vs the 36% baseline, and Pie #2 + Pie #3.

It is an ORCHESTRATOR. The heavy lifting (overnight routing, salaried/training
exclusion, proportional BOH station explosion, scheduled-vs-actual join) stays in
the canonical `build_payroll_with_scheduled.py`, which this app imports. The pies
are rendered by the canonical `render_pies.py`. This app owns only the glue,
the BAU/usage/COGS math, and two fixes that the raw engine doesn't have yet:

  1. ROBUST SHIFTS PARSER. The Sling export's block order is not stable. Some
     weeks it is [name, time+duration, role]; other weeks (incl. the one this was
     built against) it is [time+duration, role-bullet-location, name]. The engine's
     parse_cell assumes the first order and silently returns 0 scheduled hours on
     the second. This app injects an order-independent parser that classifies each
     line by content (duration vs role vs name), so it works either way.
     >> This same fix should be committed into build_payroll_with_scheduled.py
        (the canonical script) in Claude Code so the engine is robust standalone.

  2. MISSING-STATION-TAG HANDLING. The RP BOH Assembly/Prep split needs Sling
     station tags (Mainline / Grill / Kitchen Prep / Dishwasher). Some weeks every
     BOH role is tagged only by location ("Happy Eatery"). When that happens the
     split is not derivable from the data, so the app reports RP BOH whole and
     marks Assembly/Prep as DATA MISSING rather than fabricating a $0 Prep bucket.

        RP BOH Assembly = SL+Other + Mainline + Grill + Dishwasher
        RP BOH Prep     = Kitchen Prep

Conventions locked to the L10 spec:
  - Operating week = Mon-Sun (auto-derived from the payroll filename, or --week).
  - Salaried $ = HEE WEEKLY SALARIES + AN WEEKLY SALARIES (BAU), not from payroll.
  - COGS headline + pie use BAU (HEE COGS $ + AN COGS $). Usage report is computed
    for reconciliation only; a gap > $50 is flagged for Tony.
  - COGS baseline = 36% of combined Rosslyn net sales.
  - Combined net sales (the denominator for all %s) = HEE + AN NEW NET ADJ SALES.

Usage
-----
    python3 l10_labor_cogs_app.py
    python3 l10_labor_cogs_app.py --inputs-dir /mnt/project --week "6/8 - 6/14"
    python3 l10_labor_cogs_app.py --no-pies        # skip chart rendering

Outputs (to --out-dir, default /mnt/user-data/outputs):
    L10_labor_cogs_<week>.md     the summary report
    L10_pie2_labor_cost_<week>.png
    L10_pie3_labor_cogs_<week>.png
"""

import argparse
import json
import os
import re
import subprocess
import sys
from datetime import date
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------- #
# Locations
# --------------------------------------------------------------------------- #
SCRIPT_DIR = Path(__file__).resolve().parent
COGS_BASELINE_PCT = 0.36
VARIANCE_FLAG_HOURS = 2.0


def find_pie_renderer(override=None):
    """Locate render_pies.py: explicit override, then next to this script,
    then the Claude skills path (only exists inside Claude's environment)."""
    candidates = [override] if override else []
    candidates += [
        SCRIPT_DIR / "render_pies.py",
        SCRIPT_DIR / "scripts" / "render_pies.py",
        Path("/mnt/skills/user/l10-prime-cost-pie-charts/scripts/render_pies.py"),
    ]
    for c in candidates:
        if c and Path(c).exists():
            return str(c)
    return None


# --------------------------------------------------------------------------- #
# Robust, order-independent shifts parser (the pattern-based fix)
# --------------------------------------------------------------------------- #
_TIME_RE = re.compile(r"\d{1,2}:\d{2}\s*[AP]M", re.I)
_DUR_TOKEN_RE = re.compile(r"\d+\s*h(?:\s*\d+\s*min)?|\d+\s*min", re.I)
_BULLET = "\u2022"


def _hours_from(line):
    """Pull a float hour count out of a duration line ('7h 30min' -> 7.5)."""
    m = re.search(r"(\d+)\s*h(?:\s*(\d+)\s*min)?", line or "", re.I)
    if m:
        return int(m.group(1)) + (int(m.group(2)) / 60.0 if m.group(2) else 0.0)
    m2 = re.search(r"(\d+)\s*min", line or "", re.I)
    return int(m2.group(1)) / 60.0 if m2 else 0.0


def robust_parse_cell(text):
    """Split one day's Sling cell into [(name, hours, role_full), ...].

    Classifies each line of a block by CONTENT, not position, so it handles both
    [name, duration, role] and [duration, role, name] block orders. `role_full`
    preserves the 'BASE <bullet> TAG' string so the engine's split_sling_role can
    extract the station/location tag downstream.
    """
    if not text:
        return []
    out = []
    for blk in re.split(r"\n\s*\n", str(text).strip()):
        lines = [ln.strip() for ln in blk.split("\n") if ln.strip()]
        if len(lines) < 2:
            continue
        dur = role = name = None
        for ln in lines:
            is_dur = bool(_TIME_RE.search(ln) or _DUR_TOKEN_RE.search(ln))
            has_sep = (_BULLET in ln) or (" . " in ln)
            if is_dur and dur is None:
                dur = ln
            elif has_sep and role is None:
                role = ln
            elif name is None:
                name = ln
            elif role is None:
                role = ln
            elif dur is None:
                dur = ln
        if name is None:                      # positional fallback
            name = lines[-1]
        if name.lower().startswith(("unassigned", "available")):
            continue
        out.append((name, _hours_from(dur), role or ""))
    return out


# --------------------------------------------------------------------------- #
# Money parsing
# --------------------------------------------------------------------------- #
def money(text):
    """'$147,811.36' / '45,806' / '40.00%' (->0.4 NOT handled here) -> float."""
    s = str(text).strip()
    if not s or s in ("-", "#DIV/0!", "nan"):
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    cleaned = re.sub(r"[^0-9.\-]", "", s)
    if cleaned in ("", "-", ".", "-."):
        return 0.0
    val = float(cleaned)
    return -val if neg else val


# --------------------------------------------------------------------------- #
# File discovery (keyword-based, tolerant of messy export names)
# --------------------------------------------------------------------------- #
# Matching is by lowercased keyword/extension, not rigid globs, so names like
# "OJB Store's BAUs (Weekly Numbers) - HEE.csv" or "payroll export.csv" all work.

def list_dir_files(roots):
    """Return [(filename, fullpath)] across all search roots (de-duped by name)."""
    seen, out = set(), []
    for root in roots:
        try:
            for p in sorted(Path(root).iterdir()):
                if p.is_file() and p.name not in seen:
                    seen.add(p.name)
                    out.append((p.name, str(p)))
        except (FileNotFoundError, NotADirectoryError):
            continue
    return out


def _pick(all_files, contains, ends=None, exclude=None):
    """Return the best full path whose lowercased name has all `contains` tokens,
    ends with one of `ends`, and has none of `exclude`. Newest-named wins on ties."""
    ends = tuple(e.lower() for e in (ends or []))
    exclude = [x.lower() for x in (exclude or [])]
    hits = []
    for name, full in all_files:
        low = name.lower()
        if all(tok in low for tok in contains) \
           and (not ends or low.endswith(ends)) \
           and not any(x in low for x in exclude):
            hits.append((name, full))
    return sorted(hits)[-1][1] if hits else None


def discover_files(roots):
    """Locate every input by keyword across the search roots."""
    af = list_dir_files(roots)

    # BAU files: identify by HEE vs a standalone 'an' token (won't catch 'numbers').
    bau = [(n, f) for (n, f) in af
           if n.lower().endswith(".csv") and ("bau" in n.lower() or "ojb" in n.lower() or "hee" in n.lower())]
    hee = _pick(bau, ["hee"], ends=[".csv"])
    an_cands = [(n, f) for (n, f) in bau if "hee" not in n.lower()]
    an = None
    if len(an_cands) == 1:
        an = an_cands[0][1]
    elif an_cands:
        an = _pick([(n, f) for (n, f) in an_cands
                    if re.search(r"(^|[^a-z])an([^a-z]|$)", n.lower())], [], ends=[".csv"]) \
             or sorted(an_cands)[-1][1]

    return {
        "payroll":   _pick(af, ["payroll"], ends=[".csv"]),
        "shifts":    _pick(af, ["shift"],   ends=[".xls", ".xlsx"]),
        "usage":     _pick(af, ["usage"],   ends=[".csv"]),
        "bau_hee":   hee,
        "bau_an":    an,
        "salaried":  _pick(af, ["salaried"], ends=[".txt"]),
        "overnight": _pick(af, ["overnight"], ends=[".txt"]),
    }, af


def derive_week(payroll_path):
    """Derive the operating week from the payroll filename's two date stamps.

    Handles separators between the dates (or none):
      payroll_export_2026_06_152026_06_21.csv     (concatenated)
      payroll_export_2026_06_15-2026_06_21.csv    (dash)
      payroll_export_2026_06_15_to_2026_06_21.csv (word)
    and date separators _ - / .  -> ('6/15 - 6/21', start_date, end_date).
    """
    name = Path(payroll_path).name
    m = re.search(
        r"(\d{4})[_\-/.](\d{2})[_\-/.](\d{2})"
        r"[\s_\-/.\u2013\u2014]*(?:to[\s_\-/.\u2013\u2014]*)?"
        r"(\d{4})[_\-/.](\d{2})[_\-/.](\d{2})",
        name, re.I,
    )
    if not m:
        return None, None, None
    y1, m1, d1, y2, m2, d2 = (int(x) for x in m.groups())
    sd, ed = date(y1, m1, d1), date(y2, m2, d2)
    label = f"{sd.month}/{sd.day} - {ed.month}/{ed.day}"
    return label, sd, ed


# --------------------------------------------------------------------------- #
# BAU extraction
# --------------------------------------------------------------------------- #
def bau_row(path, start_date):
    """Return the BAU row whose START DATE matches the operating-week Monday.

    Tolerant of leading zeros and 2- vs 4-digit years: '6/15/26', '06/15/2026',
    '6/15/2026' all match the same Monday.
    """
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    date_col = resolve_col(df.iloc[0] if len(df) else None, "START DATE") or "START DATE"
    if date_col not in df.columns:
        raise ValueError(f"{path} has no 'START DATE' column")

    def norm_date(s):
        m = re.match(r"\s*(\d{1,2})/(\d{1,2})/(\d{2,4})\s*$", str(s))
        if not m:
            return None
        mo, da, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return (mo, da, yr % 100)

    want = (start_date.month, start_date.day, start_date.year % 100)
    norm = df[date_col].map(norm_date)
    hit = df[norm == want]
    return hit.iloc[0] if not hit.empty else None


def _norm_header(s):
    """Lowercase + collapse whitespace, for tolerant column matching."""
    return re.sub(r"\s+", " ", str(s).strip()).lower()


def resolve_col(row, col):
    """Find a BAU column by name, tolerant of header whitespace/case drift.

    Used so 'NEW NET ADJ SALES' (and the other BAU fields) are located even if
    the exported header is 'New Net Adj Sales', has trailing spaces, etc.
    Returns the actual column label present in the row, or None.
    """
    if row is None:
        return None
    if col in row.index:
        return col
    target = _norm_header(col)
    for c in row.index:
        if _norm_header(c) == target:
            return c
    return None


def get(row, col):
    """Read a BAU column value as a float, via tolerant header matching."""
    actual = resolve_col(row, col)
    return money(row[actual]) if actual is not None else 0.0


# --------------------------------------------------------------------------- #
# Labor engine
# --------------------------------------------------------------------------- #
def run_engine(files):
    """Import the canonical engine, inject the robust parser, run build()."""
    for d in (str(SCRIPT_DIR), str(Path(files["payroll"]).parent), "/mnt/project"):
        if d and d not in sys.path:
            sys.path.insert(0, d)
    try:
        import build_payroll_with_scheduled as eng
    except ImportError:
        sys.stderr.write(
            "ERROR: could not import build_payroll_with_scheduled.py.\n"
            f"       Put that file next to this app ({SCRIPT_DIR}) or in the inputs folder.\n"
        )
        sys.exit(1)

    eng.parse_cell = robust_parse_cell                # <-- the fix
    detail = eng.build(
        payroll_path=files["payroll"],
        shifts_path=files["shifts"],
        salaried_path=files["salaried"] or str(SCRIPT_DIR / "SALARIED_EMPLOYEES.txt"),
        overnight_path=files["overnight"] or str(SCRIPT_DIR / "Overnight_Employees.txt"),
    )
    return eng, detail


def summarize_labor(eng, detail):
    """Roll the engine detail up into the buckets the L10 needs."""
    detail = detail.copy()
    detail["_parent"] = detail["_bucket"].apply(eng.parent_bucket)
    detail["_act_h"] = detail["Regular Hours"] + detail["Overtime Hours"]

    def pay(mask):
        return round(detail.loc[mask, "Total Pay"].sum(), 2)

    def hrs(mask, col="Scheduled Hours"):
        return round(detail.loc[mask, col].sum(), 2)

    an_foh = pay(detail["_parent"] == "AN FOH")
    an_boh = pay(detail["_parent"] == "AN BOH")
    rp_foh = pay(detail["_parent"] == "RP FOH")
    rp_boh = pay(detail["_parent"] == "RP BOH")

    stations = {st: pay(detail["_bucket"] == st) for st in eng.RP_BOH_STATIONS}
    # Assembly/Prep only if real station tags exist (non-SL stations carry $).
    tagged = sum(v for k, v in stations.items() if k != "RP BOH - SL + Other")
    split_available = round(tagged, 2) > 0
    assembly = round(
        stations["RP BOH - SL + Other"] + stations["RP BOH - Mainline"]
        + stations["RP BOH - Grill"] + stations["RP BOH - Dishwasher"], 2)
    prep = stations["RP BOH - Kitchen Prep"]

    # Scheduled vs Actual (hourly only; salaried + training already excluded by engine)
    sched_dollars = round(detail["Expected Pay"].sum(), 2)
    actual_dollars = round(detail["Total Pay"].sum(), 2)
    sched_hours = round(detail["Scheduled Hours"].sum(), 2)
    actual_hours = round(detail["_act_h"].sum(), 2)
    ot_hours = round(detail["Overtime Hours"].sum(), 2)

    emp = detail.groupby("Employee").agg(
        sched=("Scheduled Hours", "sum"),
        actual=("_act_h", "sum"),
        ot=("Overtime Hours", "sum"),
        pay=("Total Pay", "sum"),
    ).reset_index()
    emp["var_h"] = (emp["actual"] - emp["sched"]).round(2)
    big = emp[emp["var_h"].abs() >= VARIANCE_FLAG_HOURS].copy()
    big = big.reindex(big["var_h"].abs().sort_values(ascending=False).index)

    return {
        "an_foh": an_foh, "an_boh": an_boh, "rp_foh": rp_foh, "rp_boh": rp_boh,
        "stations": stations, "split_available": split_available,
        "rp_boh_assembly": assembly, "rp_boh_prep": prep,
        "hourly_total": round(an_foh + an_boh + rp_foh + rp_boh, 2),
        "sched_dollars": sched_dollars, "actual_dollars": actual_dollars,
        "sched_hours": sched_hours, "actual_hours": actual_hours, "ot_hours": ot_hours,
        "big_variance": big,
    }


# --------------------------------------------------------------------------- #
# COGS
# --------------------------------------------------------------------------- #
def usage_cogs(path):
    u = pd.read_csv(path, encoding="utf-8-sig")
    food = u.loc[u["Category Type"] == "Food", "Used Value"].sum()
    nabev = u.loc[u["Category Type"] == "N/A Bev", "Used Value"].sum()
    paper = u.loc[(u["Category Type"] == "Other")
                  & (u["Category"] == "Paper and Packaging"), "Used Value"].sum()
    return {"food": round(food, 2), "nabev": round(nabev, 2),
            "paper": round(paper, 2), "total": round(food + nabev + paper, 2)}


# --------------------------------------------------------------------------- #
# Pie rendering (handoff to the canonical renderer)
# --------------------------------------------------------------------------- #
def render_pies(week, pie_inputs, out_dir, renderer):
    if not renderer:
        sys.stderr.write("WARN: render_pies.py not found (looked next to the app and "
                         "in ./scripts). Copy it locally or pass --pie-renderer; skipping pies.\n")
        return {}
    ipath = Path(out_dir) / "_pie_inputs.json"
    ipath.write_text(json.dumps(pie_inputs, indent=2), encoding="utf-8")
    res = subprocess.run(
        [sys.executable, renderer, "--week", week.replace("-", "\u2013"),
         "--inputs", str(ipath), "--pies", "2,3", "--out-dir", str(out_dir)],
        capture_output=True, text=True,
    )
    sys.stderr.write(res.stderr)
    try:
        return json.loads(res.stdout)
    except json.JSONDecodeError:
        return {}


# --------------------------------------------------------------------------- #
# Report
# --------------------------------------------------------------------------- #
def fmt(x):
    return f"${x:,.2f}"


def pct(part, whole):
    return f"{part / whole * 100:.2f}%" if whole else "n/a"


def build_report(ctx):
    L = ctx["labor"]
    s = ctx
    lines = []
    P = lines.append

    P(f"# L10 — Labor & COGs — Operating Week {s['week']}")
    P("")
    P(f"**L10 meeting date:** Mon after week close  |  **Owner:** Donovan")
    P(f"**Sources:** payroll `{Path(s['files']['payroll']).name}` • "
      f"shifts `{Path(s['files']['shifts']).name}` • usage `{Path(s['files']['usage']).name}` • "
      f"BAU HEE/AN • Salaried/Overnight rosters")
    P("")
    P("**Data integrity notes:**")
    for n in s["flags"]:
        P(f"- {n}")
    P("")
    P("---")
    P("## Sales (denominator)")
    P(f"- Josh Kim — AN NEW ADJ = **{fmt(s['an_sales'])}**")
    P(f"- David Nguyen — RP NEW ADJ = **{fmt(s['rp_sales'])}**")
    P(f"- **RP + AN Combined Net Sales = {fmt(s['net_sales'])}**")
    P("")
    P("---")
    P("## Labor")
    tot = s["total_labor"]
    P(f"- **Total Labor = {fmt(tot)} ({pct(tot, s['net_sales'])} of sales)**")
    P(f"  - AN = {fmt(L['an_foh'] + L['an_boh'])} ({pct(L['an_foh'] + L['an_boh'], s['net_sales'])})")
    P(f"    - FOH = {fmt(L['an_foh'])} ({pct(L['an_foh'], s['net_sales'])})")
    P(f"    - BOH = {fmt(L['an_boh'])} ({pct(L['an_boh'], s['net_sales'])})")
    P(f"  - RP = {fmt(L['rp_foh'] + L['rp_boh'])} ({pct(L['rp_foh'] + L['rp_boh'], s['net_sales'])})")
    P(f"    - FOH = {fmt(L['rp_foh'])} ({pct(L['rp_foh'], s['net_sales'])})")
    P(f"    - BOH = {fmt(L['rp_boh'])} ({pct(L['rp_boh'], s['net_sales'])})")
    if L["split_available"]:
        P(f"      - Assembly (SL+Line+Grill+Dish) = {fmt(L['rp_boh_assembly'])} ({pct(L['rp_boh_assembly'], s['net_sales'])})")
        P(f"      - Prep (Kitchen Prep) = {fmt(L['rp_boh_prep'])} ({pct(L['rp_boh_prep'], s['net_sales'])})")
    else:
        P("      - Assembly vs Prep: **DATA MISSING** — shifts export has no BOH station "
          "tags this week (all roles tagged by location). RP BOH reported whole.")
    P(f"  - Salaried = {fmt(s['salaried'])} ({pct(s['salaried'], s['net_sales'])})  "
      f"[HEE {fmt(s['salaried_hee'])} + AN {fmt(s['salaried_an'])}]")
    P("")
    P("---")
    P("## Labor — Scheduled vs Actual (hourly; excludes salaried + training)")
    var_d = L["actual_dollars"] - L["sched_dollars"]
    var_h = L["actual_hours"] - L["sched_hours"]
    P(f"- Scheduled = {L['sched_hours']:.1f}h → {fmt(L['sched_dollars'])}")
    P(f"- Actual = {L['actual_hours']:.1f}h → {fmt(L['actual_dollars'])}")
    P(f"- **Variance = {var_d:+,.2f} ({var_d / L['sched_dollars'] * 100:+.1f}%) / {var_h:+.1f}h** "
      f"(OT this week: {L['ot_hours']:.1f}h)")
    P("")
    P(f"- Employees over/under by ≥ {VARIANCE_FLAG_HOURS:g}h:")
    if len(L["big_variance"]):
        P("")
        P("  | Employee | Sched h | Actual h | Var h | OT h | Pay |")
        P("  |---|--:|--:|--:|--:|--:|")
        for _, r in L["big_variance"].iterrows():
            P(f"  | {r['Employee']} | {r['sched']:.1f} | {r['actual']:.1f} | "
              f"{r['var_h']:+.1f} | {r['ot']:.1f} | {fmt(r['pay'])} |")
    else:
        P("  - none")
    P("")
    P("---")
    P("## COGS")
    cb = s["cogs_bau"]
    base = s["net_sales"] * COGS_BASELINE_PCT
    P(f"- **Weekly Total (BAU) = {fmt(cb)} ({pct(cb, s['net_sales'])})**  "
      f"[HEE {fmt(s['cogs_hee'])} + AN {fmt(s['cogs_an'])}]")
    P(f"- Baseline = {fmt(base)} (36%)")
    P(f"- **Variance vs baseline = {cb - base:+,.2f} ({(cb / s['net_sales'] * 100 - 36):+.2f} pp)**")
    P(f"- Usage-report cross-check = {fmt(s['cogs_usage']['total'])} "
      f"(Food {fmt(s['cogs_usage']['food'])} + N/A Bev {fmt(s['cogs_usage']['nabev'])} "
      f"+ Paper&Pkg {fmt(s['cogs_usage']['paper'])})")
    gap = s["cogs_usage"]["total"] - cb
    P(f"- Usage vs BAU gap = {gap:+,.2f} → **{'flag for Tony' if abs(gap) > 50 else 'reconciles'}**; "
      f"BAU used for headline + pie")
    P("")
    P("---")
    P("## Prime Cost")
    prime = s["total_labor"] + cb
    P(f"- Total Labor {fmt(s['total_labor'])} + COGS {fmt(cb)} = "
      f"**Prime Cost {fmt(prime)} ({pct(prime, s['net_sales'])})**")
    P(f"- Gross margin = {fmt(s['net_sales'] - prime)} ({pct(s['net_sales'] - prime, s['net_sales'])})")
    P("")
    if s.get("pie_files"):
        P("---")
        P("## Pie Charts")
        for f in s["pie_files"]:
            P(f"- `{Path(f).name}`")
    return "\n".join(lines)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    # Windows consoles default to cp1252, which can't encode '->'/'>=' glyphs in
    # the report. Make console output UTF-8 (replace, never crash).
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--inputs-dir", default=str(SCRIPT_DIR),
                    help="Folder holding the week's files (default: the app's own folder)")
    ap.add_argument("--out-dir", default=str(SCRIPT_DIR / "output"),
                    help="Where to write the report + pies (default: ./output)")
    ap.add_argument("--week", default=None, help='Override, e.g. "6/8 - 6/14"')
    ap.add_argument("--pie-renderer", default=None, help="Path to render_pies.py")
    ap.add_argument("--no-pies", action="store_true")
    args = ap.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Search the inputs folder AND the app's own folder (so "kept on hand" files —
    # rosters, the engine, the pie renderer — can live next to the app).
    roots = list(dict.fromkeys([args.inputs_dir, str(SCRIPT_DIR)]))
    files, all_files = discover_files(roots)
    flags = []

    required = ["payroll", "shifts", "usage", "bau_hee", "bau_an"]   # overnight/salaried optional
    missing = [k for k in required if files.get(k) is None]
    if missing:
        sys.stderr.write(f"\nERROR: could not locate required input(s): {missing}\n")
        sys.stderr.write(f"Searched: {roots}\n")
        sys.stderr.write("Files I can actually see in those folders:\n")
        for name, _ in all_files:
            sys.stderr.write(f"   - {name}\n")
        if not all_files:
            sys.stderr.write("   (none — the folder is empty or the path is wrong)\n")
        sys.stderr.write("\nFix: put the week's files in the app's folder, or run with "
                         "--inputs-dir \"C:\\path\\to\\your\\files\"\n")
        sys.exit(1)
    for opt in ("salaried", "overnight"):
        if files.get(opt) is None:
            files[opt] = None
            sys.stderr.write(f"WARN: {opt} roster not found; continuing "
                             f"({'overnight routing degrades' if opt == 'overnight' else 'no effect on totals'}).\n")

    sys.stderr.write("Resolved inputs:\n")
    for k in ["payroll", "shifts", "usage", "bau_hee", "bau_an", "salaried", "overnight"]:
        sys.stderr.write(f"   {k:9} -> {files[k]}\n")

    week, sd, ed = derive_week(files["payroll"])
    if args.week:
        week = args.week
    if sd is None and not args.week:
        sys.stderr.write("ERROR: could not derive the week from the payroll filename; pass --week\n")
        sys.exit(1)

    # ---- Sales + Salaried + COGS from BAU ----
    hee = bau_row(files["bau_hee"], sd) if sd else None
    an = bau_row(files["bau_an"], sd) if sd else None
    if hee is None:
        flags.append(f"DATA MISSING — HEE BAU has no row for week start {sd}")
    if an is None:
        flags.append(f"DATA MISSING — AN BAU has no row for week start {sd}")

    # ---- Sales come from the BAU 'NEW NET ADJ SALES' columns (RP=HEE, AN=AN) ----
    SALES_COL = "NEW NET ADJ SALES"
    rp_sales = get(hee, SALES_COL)
    an_sales = get(an, SALES_COL)
    for label, row in (("HEE", hee), ("AN", an)):
        if row is not None and resolve_col(row, SALES_COL) is None:
            flags.append(f"WARNING — '{SALES_COL}' column not found in {label} BAU; "
                         f"{label} sales read as $0. Check the BAU header.")
            sys.stderr.write(f"WARN: '{SALES_COL}' not found in {label} BAU header.\n")
    if rp_sales == 0 or an_sales == 0:
        sys.stderr.write(f"WARN: a NEW NET ADJ SALES value is $0 (RP={rp_sales}, AN={an_sales}); "
                         f"confirm the BAU row matched the week start {sd}.\n")
    net_sales = rp_sales + an_sales
    salaried_hee = get(hee, "WEEKLY SALARIES")
    salaried_an = get(an, "WEEKLY SALARIES")
    salaried = salaried_hee + salaried_an
    cogs_hee = get(hee, "COGS $")
    cogs_an = get(an, "COGS $")
    cogs_bau = cogs_hee + cogs_an
    cogs_use = usage_cogs(files["usage"])

    # ---- Labor engine ----
    eng, detail = run_engine(files)
    labor = summarize_labor(eng, detail)
    total_labor = round(labor["hourly_total"] + salaried, 2)

    # ---- Flags ----
    if not labor["split_available"]:
        flags.append("RP BOH Assembly/Prep split unavailable — shifts export carries no "
                     "BOH station tags this week (roles tagged by location only).")
    if abs(cogs_use["total"] - cogs_bau) > 50:
        flags.append(f"COGS source conflict — usage report {fmt(cogs_use['total'])} vs "
                     f"BAU {fmt(cogs_bau)} (gap {cogs_use['total'] - cogs_bau:+,.2f}); "
                     f"BAU used. Reconcile with Tony.")
    if an is not None and abs(get(an, 'COGS %') - 40.0) < 0.01:
        flags.append("AN COGS posts at a flat 40.00% of sales (formula-driven, not actuals). "
                     "Confirm with Tony before the meeting.")
    flags.append("Shifts parsed with order-independent parser (engine's parse_cell would "
                 "have returned 0 scheduled hours on this week's block order).")
    flags.append("AN BOH hourly = engine bucket (An Line Lead is classified AN FOH, which "
                 "reconciles to BAU 'AN FOH HOURLY WAGES'); AN's BOH labor is salaried.")

    # ---- Pies ----
    pie_files = []
    pie_inputs = {
        "an_foh": labor["an_foh"], "rp_foh": labor["rp_foh"],
        "an_boh": labor["an_boh"], "rp_boh": labor["rp_boh"],
        "salaried": salaried, "cogs": cogs_bau,
        "net_sales": net_sales, "cogs_source": "BAU",
    }
    if not args.no_pies:
        renderer = find_pie_renderer(args.pie_renderer)
        summary = render_pies(week, pie_inputs, out_dir, renderer)
        pie_files = summary.get("files", [])

    ctx = {
        "week": week, "files": files, "flags": flags,
        "rp_sales": rp_sales, "an_sales": an_sales, "net_sales": net_sales,
        "salaried": salaried, "salaried_hee": salaried_hee, "salaried_an": salaried_an,
        "cogs_bau": cogs_bau, "cogs_hee": cogs_hee, "cogs_an": cogs_an,
        "cogs_usage": cogs_use, "labor": labor, "total_labor": total_labor,
        "pie_files": pie_files,
    }
    report = build_report(ctx)

    report_path = out_dir / f"L10_labor_cogs_{week.replace('/', '_').replace(' ', '').replace('-', '_')}.md"
    report_path.write_text(report, encoding="utf-8")

    print(report)
    print(f"\n[wrote {report_path}]", file=sys.stderr)
    for f in pie_files:
        print(f"[wrote {f}]", file=sys.stderr)


if __name__ == "__main__":
    main()
