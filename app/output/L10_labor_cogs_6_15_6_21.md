# L10 — Labor & COGs — Operating Week 6/15 - 6/21

**L10 meeting date:** Mon after week close  |  **Owner:** Donovan
**Sources:** payroll `payroll_export_2026_06_15-2026_06_21.csv` • shifts `shifts-export.xls` • usage `usageReport.csv` • BAU HEE/AN • Salaried/Overnight rosters

**Data integrity notes:**
- AN COGS posts at a flat 40.00% of sales (formula-driven, not actuals). Confirm with Tony before the meeting.
- Shifts parsed with order-independent parser (engine's parse_cell would have returned 0 scheduled hours on this week's block order).
- AN BOH hourly = engine bucket (An Line Lead is classified AN FOH, which reconciles to BAU 'AN FOH HOURLY WAGES'); AN's BOH labor is salaried.

---
## Sales (denominator)
- Josh Kim — AN NEW ADJ = **$16,762.22**
- David Nguyen — RP NEW ADJ = **$135,938.35**
- **RP + AN Combined Net Sales = $152,700.57**

---
## Labor
- **Total Labor = $46,106.03 (30.19% of sales)**
  - AN = $4,788.08 (3.14%)
    - FOH = $3,259.73 (2.13%)
    - BOH = $1,528.35 (1.00%)
  - RP = $28,420.95 (18.61%)
    - FOH = $6,832.41 (4.47%)
    - BOH = $21,588.54 (14.14%)
      - Assembly (SL+Line+Grill+Dish) = $17,912.73 (11.73%)
      - Prep (Kitchen Prep) = $3,675.81 (2.41%)
  - Salaried = $12,897.00 (8.45%)  [HEE $11,628.00 + AN $1,269.00]

---
## Labor — Scheduled vs Actual (hourly; excludes salaried + training)
- Scheduled = 1666.8h → $32,562.74
- Actual = 1674.0h → $33,209.03
- **Variance = +646.29 (+2.0%) / +7.3h** (OT this week: 24.7h)

- Employees over/under by ≥ 2h:

  | Employee | Sched h | Actual h | Var h | OT h | Pay |
  |---|--:|--:|--:|--:|--:|
  | Woods, Nya | 26.0 | 37.9 | +11.9 | 0.0 | $624.97 |
  | Veizaga Zelada, Laura | 22.0 | 13.1 | -8.9 | 0.0 | $223.10 |
  | Alcantara, Salvador | 36.5 | 45.0 | +8.5 | 5.0 | $841.18 |
  | Hinojosa Orosco, Josselyn | 32.0 | 40.3 | +8.3 | 0.3 | $802.95 |
  | Vasquez, Jessy | 31.2 | 23.2 | -8.0 | 0.0 | $418.14 |
  | Severichs, Andree | 38.0 | 31.2 | -6.8 | 0.0 | $576.07 |
  | Alvarado, Orlando | 35.0 | 28.9 | -6.1 | 0.0 | $741.57 |
  | Dridi, Amina | 19.2 | 13.4 | -5.9 | 0.0 | $233.87 |
  | Rivera Guevara, Felix | 45.5 | 40.2 | -5.3 | 0.2 | $988.57 |
  | Montano, Marina | 48.0 | 43.1 | -4.9 | 3.1 | $768.62 |
  | Veizaga, Panfilo | 34.0 | 37.8 | +3.8 | 0.0 | $692.90 |
  | Almendras, Jimena | 26.5 | 30.0 | +3.5 | 0.0 | $523.72 |
  | Castellon Jimenez, Beimar | 40.5 | 37.1 | -3.4 | 0.0 | $1,109.62 |
  | Ugarte, Jhonn | 17.0 | 13.9 | -3.1 | 0.0 | $246.02 |
  | Belay, Daniel | 11.8 | 14.3 | +2.5 | 0.0 | $221.67 |
  | Clayborne, Tayy | 30.5 | 33.0 | +2.5 | 0.0 | $792.82 |
  | Mamani, Sabino | 39.0 | 41.5 | +2.5 | 1.5 | $753.20 |
  | Wahyu Dyatmika, Komang Nando | 22.0 | 19.6 | -2.4 | 0.0 | $327.28 |
  | Fernandez, Fredy | 38.0 | 40.3 | +2.3 | 0.3 | $761.62 |
  | Ramos, Beatriz | 28.0 | 30.1 | +2.1 | 0.0 | $516.91 |

---
## COGS
- **Weekly Total (BAU) = $58,743.00 (38.47%)**  [HEE $52,038.00 + AN $6,705.00]
- Baseline = $54,972.21 (36%)
- **Variance vs baseline = +3,770.79 (+2.47 pp)**
- Usage-report cross-check = $58,743.10 (Food $51,319.78 + N/A Bev $1,186.48 + Paper&Pkg $6,236.85)
- Usage vs BAU gap = +0.10 → **reconciles**; BAU used for headline + pie

---
## Prime Cost
- Total Labor $46,106.03 + COGS $58,743.00 = **Prime Cost $104,849.03 (68.66%)**
- Gross margin = $47,851.54 (31.34%)

---
## Pie Charts
- `L10_pie2_labor_cost_6_15_6_21.png`
- `L10_pie3_labor_cogs_6_15_6_21.png`