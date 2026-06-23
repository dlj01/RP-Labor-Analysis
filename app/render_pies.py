#!/usr/bin/env python3
"""
render_l10_pies.py — Generate the L10 prime-cost pie charts (Pie #1, #2, #3).

Matplotlib renderer encoding the locked L10 pie spec:
  - Fixed slice -> color palette
  - Labels show "$ amount (%)"; black text, white fill, border matching the slice
  - Title + subtitle in a single bordered box
  - Prime Cost callout in highlighter-yellow with a black border (Pies #1 and #3)
  - Slice 1 starts at 12 o'clock and sweeps CLOCKWISE (per spec)
  - Filenames use underscores only

Called by l10_labor_cogs_app.py, or run standalone:
    python render_l10_pies.py --week "6/15 - 6/21" --inputs inputs.json --pies 2,3
    python render_l10_pies.py --week "6/15 - 6/21" --inputs inputs.json --pies 1,2,3 --counterclockwise

Inputs JSON shape:
    {
      "an_foh": 3375.72, "rp_foh": 7151.77, "an_boh": 0.0, "rp_boh": 21954.10,
      "salaried": 12897.00, "cogs": 53128.00, "net_sales": 166115.15,
      "cogs_source": "BAU"
    }

stdout: a single JSON object {"files": [...], ...} (the app reads this).
stderr: progress / warnings.
"""

import argparse
import json
import sys
from pathlib import Path

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np
except ImportError:
    sys.stderr.write("ERROR: matplotlib and numpy are required (pip install matplotlib numpy)\n")
    sys.exit(2)

# --------------------------------------------------------------------------- #
# Locked palette
# --------------------------------------------------------------------------- #
C_AN_FOH = "#F4D03F"      # yellow
C_RP_FOH = "#E74C3C"      # red
C_AN_BOH = "#27AE60"      # green
C_RP_BOH = "#3498DB"      # blue
C_SALARIED = "#8E44AD"    # purple
C_COGS_BROWN = "#8B4513"  # brown  (Pie #1)
C_COGS_TEAL = "#16A085"   # teal   (Pie #3)
C_LABOR = "#E67E22"       # orange (Pie #3)
C_GROSS = "#BDC3C7"       # gray
C_PRIME_BG = "#FFFF66"    # highlighter yellow
C_PRIME_BORDER = "#000000"

# Spec: slice 1 starts at 12 o'clock and goes clockwise -> counterclock=False.
# Set to True (or pass --counterclockwise) to match pre-spec rendered weeks.
COUNTERCLOCK = False


def make_label_box(ax, x, y, lines, border_color, bg="#FFFFFF", fontsize=11):
    bbox = dict(boxstyle="round,pad=0.4,rounding_size=0.3",
                facecolor=bg, edgecolor=border_color, linewidth=2)
    ax.text(x, y, "\n".join(lines), ha="center", va="center",
            fontsize=fontsize, color="black", bbox=bbox, zorder=10)


def radius_factor(pct):
    """Small slices push the label outward; large slices keep it inside."""
    if pct >= 25:
        return 0.55
    elif pct >= 10:
        return 0.90
    return 1.15


def place_label(ax, wedge, label, val, color, denom, fontsize=11):
    pct = (val / denom) * 100 if denom > 0 else 0
    r = radius_factor(pct)
    ang = (wedge.theta1 + wedge.theta2) / 2
    x = r * np.cos(np.deg2rad(ang))
    y = r * np.sin(np.deg2rad(ang))
    make_label_box(ax, x, y, [label, f"${val:,.2f} ({pct:.1f}%)"], color, fontsize=fontsize)
    if pct < 10 and r > 1.0:                       # leader line for tiny outside labels
        ax.plot([0.85 * np.cos(np.deg2rad(ang)), x * 0.93],
                [0.85 * np.sin(np.deg2rad(ang)), y * 0.93],
                color="#666", linewidth=1, zorder=5)


def place_labels_smart(ax, wedges, slices, denom, fontsize=11, inside_cut=10.0,
                       x_right=1.72, x_left=-1.72, gap=0.36):
    """Inside labels for big slices; small slices get stacked outside columns
    (left/right) with leader lines, so adjacent thin wedges never overlap."""
    right, left = [], []
    for w, (label, val, color) in zip(wedges, slices):
        pct = val / denom * 100 if denom > 0 else 0
        ang = (w.theta1 + w.theta2) / 2
        if pct >= inside_cut:
            r = radius_factor(pct)
            make_label_box(ax, r * np.cos(np.deg2rad(ang)), r * np.sin(np.deg2rad(ang)),
                           [label, f"${val:,.2f} ({pct:.1f}%)"], color, fontsize=fontsize)
        else:
            ex, ey = np.cos(np.deg2rad(ang)), np.sin(np.deg2rad(ang))
            (right if ex >= 0 else left).append(
                dict(label=label, val=val, color=color, pct=pct, ex=ex, ey=ey))
    for items, xlab in ((right, x_right), (left, x_left)):
        if not items:
            continue
        items.sort(key=lambda d: d["ey"], reverse=True)
        ys = [it["ey"] * 1.18 for it in items]
        for i in range(1, len(ys)):                 # enforce min vertical gap top-down
            if ys[i] > ys[i - 1] - gap:
                ys[i] = ys[i - 1] - gap
        for it, yl in zip(items, ys):
            make_label_box(ax, xlab, yl,
                           [it["label"], f"${it['val']:,.2f} ({it['pct']:.1f}%)"],
                           it["color"], fontsize=fontsize)
            ax.plot([it["ex"] * 0.98, xlab], [it["ey"] * 0.98, yl],
                    color="#666", linewidth=1, zorder=5)


def week_to_filename_token(week):
    """'6/15 - 6/21' / '6/15 \u2013 6/21' -> '6_15_6_21'."""
    return (week.replace("/", "_").replace("\u2013", "_")
                .replace("-", "_").replace(" ", ""))


def _pie(ax, values, colors):
    wedges, _ = ax.pie(values, colors=colors, startangle=90,
                       counterclock=COUNTERCLOCK,
                       wedgeprops=dict(edgecolor="white", linewidth=2))
    return wedges


def render_pie1(week, inputs, out_dir):
    """Pie #1: full labor + COGS + gross margin vs net sales; Prime Cost = slices 1-6."""
    an_foh, rp_foh, an_boh, rp_boh = inputs["an_foh"], inputs["rp_foh"], inputs["an_boh"], inputs["rp_boh"]
    salaried, cogs, net_sales = inputs["salaried"], inputs["cogs"], inputs["net_sales"]
    total_labor = an_foh + rp_foh + an_boh + rp_boh + salaried
    prime_cost = total_labor + cogs
    gross_margin = net_sales - prime_cost

    slices = [("AN FOH", an_foh, C_AN_FOH), ("RP FOH", rp_foh, C_RP_FOH),
              ("AN BOH", an_boh, C_AN_BOH), ("RP BOH", rp_boh, C_RP_BOH),
              ("Salaried", salaried, C_SALARIED), ("COGS", cogs, C_COGS_BROWN),
              ("Gross Margin", gross_margin, C_GROSS)]

    fig, ax = plt.subplots(figsize=(13, 10), dpi=200)
    wedges = _pie(ax, [s[1] for s in slices], [s[2] for s in slices])
    place_labels_smart(ax, wedges, slices, net_sales)

    make_label_box(ax, -1.78, -0.55,
                   ["Prime Cost", f"${prime_cost:,.2f} ({prime_cost/net_sales*100:.1f}%)"],
                   border_color=C_PRIME_BORDER, bg=C_PRIME_BG, fontsize=12)
    make_label_box(ax, 0, 1.62,
                   [f"Labor & COGs {week}", f"RP + An net sales = ${net_sales:,.2f}"],
                   border_color="black", fontsize=13)

    ax.set_xlim(-2.4, 2.5); ax.set_ylim(-1.5, 1.95)
    ax.set_aspect("equal"); ax.axis("off")
    fname = out_dir / f"L10_pie1_labor_cogs_{week_to_filename_token(week)}.png"
    plt.savefig(str(fname), bbox_inches="tight", dpi=200, facecolor="white")
    plt.close(fig)
    return fname


def render_pie2(week, inputs, out_dir):
    """Pie #2: labor cost composition only (no COGS, no bracket)."""
    an_foh, rp_foh, an_boh, rp_boh = inputs["an_foh"], inputs["rp_foh"], inputs["an_boh"], inputs["rp_boh"]
    salaried = inputs["salaried"]
    total_labor = an_foh + rp_foh + an_boh + rp_boh + salaried

    slices = [("AN FOH", an_foh, C_AN_FOH), ("RP FOH", rp_foh, C_RP_FOH),
              ("AN BOH", an_boh, C_AN_BOH), ("RP BOH", rp_boh, C_RP_BOH),
              ("Salaried", salaried, C_SALARIED)]

    fig, ax = plt.subplots(figsize=(11, 9), dpi=200)
    wedges = _pie(ax, [s[1] for s in slices], [s[2] for s in slices])
    for w, (label, val, color) in zip(wedges, slices):
        place_label(ax, w, label, val, color, total_labor)

    make_label_box(ax, 0, 1.45,
                   [f"Labor Cost (RP + An) {week}", f"Total Labor = ${total_labor:,.2f}"],
                   border_color="black", fontsize=13)

    ax.set_xlim(-1.6, 1.6); ax.set_ylim(-1.4, 1.8)
    ax.set_aspect("equal"); ax.axis("off")
    fname = out_dir / f"L10_pie2_labor_cost_{week_to_filename_token(week)}.png"
    plt.savefig(str(fname), bbox_inches="tight", dpi=200, facecolor="white")
    plt.close(fig)
    return fname


def render_pie3(week, inputs, out_dir):
    """Pie #3: total labor + COGS + gross margin vs net sales; Prime Cost = slices 1-2."""
    an_foh, rp_foh, an_boh, rp_boh = inputs["an_foh"], inputs["rp_foh"], inputs["an_boh"], inputs["rp_boh"]
    salaried, cogs, net_sales = inputs["salaried"], inputs["cogs"], inputs["net_sales"]
    total_labor = an_foh + rp_foh + an_boh + rp_boh + salaried
    prime_cost = total_labor + cogs
    gross_margin = net_sales - prime_cost

    slices = [("Total Labor", total_labor, C_LABOR), ("COGS", cogs, C_COGS_TEAL),
              ("Gross Margin", gross_margin, C_GROSS)]

    fig, ax = plt.subplots(figsize=(11, 9), dpi=200)
    wedges = _pie(ax, [s[1] for s in slices], [s[2] for s in slices])
    for w, (label, val, color) in zip(wedges, slices):
        place_label(ax, w, label, val, color, net_sales, fontsize=12)

    make_label_box(ax, -1.35, -0.50,
                   ["Prime Cost", f"${prime_cost:,.2f} ({prime_cost/net_sales*100:.1f}%)"],
                   border_color=C_PRIME_BORDER, bg=C_PRIME_BG, fontsize=12)
    make_label_box(ax, 0, 1.45,
                   [f"Labor & COGs (RP + An) {week}", f"RP + An net sales = ${net_sales:,.2f}"],
                   border_color="black", fontsize=13)

    ax.set_xlim(-1.8, 1.6); ax.set_ylim(-1.4, 1.8)
    ax.set_aspect("equal"); ax.axis("off")
    fname = out_dir / f"L10_pie3_labor_cogs_{week_to_filename_token(week)}.png"
    plt.savefig(str(fname), bbox_inches="tight", dpi=200, facecolor="white")
    plt.close(fig)
    return fname


def main():
    global COUNTERCLOCK
    for _stream in (sys.stdout, sys.stderr):          # Windows cp1252 safety
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--week", required=True, help='Operating week, e.g. "6/15 - 6/21"')
    ap.add_argument("--inputs", required=True, help="Path to the inputs JSON file")
    ap.add_argument("--pies", default="2,3", help="Comma-separated subset of {1,2,3}. Default: 2,3")
    ap.add_argument("--out-dir", default="output", help="Output directory (default: ./output)")
    ap.add_argument("--counterclockwise", action="store_true",
                    help="Sweep counter-clockwise (matches pre-spec rendered weeks)")
    args = ap.parse_args()

    if args.counterclockwise:
        COUNTERCLOCK = True

    inputs = json.loads(Path(args.inputs).read_text(encoding="utf-8"))
    required = ["an_foh", "rp_foh", "an_boh", "rp_boh", "salaried", "cogs", "net_sales"]
    missing = [k for k in required if k not in inputs]
    if missing:
        sys.stderr.write(f"ERROR: inputs JSON missing keys: {missing}\n")
        sys.exit(1)

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    renderers = {"1": render_pie1, "2": render_pie2, "3": render_pie3}
    produced = []
    for p in sorted({x.strip() for x in args.pies.split(",")}):
        if p not in renderers:
            sys.stderr.write(f"WARN: unknown pie '{p}'; valid: 1, 2, 3\n")
            continue
        fname = renderers[p](args.week, inputs, out_dir)
        produced.append(fname)
        sys.stderr.write(f"wrote {fname}\n")

    total_labor = sum(inputs[k] for k in ["an_foh", "rp_foh", "an_boh", "rp_boh", "salaried"])
    prime_cost = total_labor + inputs["cogs"]
    net = inputs["net_sales"]
    summary = {
        "week": args.week,
        "cogs_source": inputs.get("cogs_source", "unspecified"),
        "total_labor": round(total_labor, 2),
        "labor_pct": round(total_labor / net * 100, 1) if net else 0,
        "cogs_pct": round(inputs["cogs"] / net * 100, 1) if net else 0,
        "prime_cost": round(prime_cost, 2),
        "prime_cost_pct": round(prime_cost / net * 100, 1) if net else 0,
        "files": [str(p) for p in produced],
    }
    print(json.dumps(summary, indent=2))           # ASCII-safe (ensure_ascii default)


if __name__ == "__main__":
    main()
