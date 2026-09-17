# Berth Waiting Time Analysis — THBKK & THLCH

Analysis of berth waiting time (ETA → ETB) for the THBKK (Bangkok) and THLCH (Laem Chabang)
terminals, covering 669 vessel calls from 2026-01-01 to 2026-09-16.

## Files

- **`Berth_Waiting_Time_Analysis_THBKK_THLCH.xlsx`** — Excel workbook with Raw Data plus
  Terminal, Service, Operator, Monthly, Wharf, Top-20, and Operational Insights sheets.
  All figures are live formulas (no hardcoded results); native charts included.
- **`Berth_Waiting_Time_Analysis_THBKK_THLCH.html`** — Self-contained interactive HTML
  dashboard (hover tooltips, light/dark mode) with the same analysis plus an executive
  summary and key findings.
- **`source_vessel_schedule.xls`** — Original uploaded vessel schedule the analysis was
  built from.

## Headline findings

- THBKK vessels wait ~2.8× longer than THLCH on average (11.2 vs 4.0 hrs), largely because
  THBKK is a river terminal behind the Bangkok Bar — vessels must wait for high tide to
  proceed upriver to berth. This shows up directly in the data as a two-peak berthing-time
  pattern roughly 12–13 hours apart (semi-diurnal tide signature), while THLCH's berthing
  times are spread flat across the day.
- Slowest berth: THBKK · BKK01. Slowest operator (OPR): HAS. Slowest service: KHS1.
- See the HTML dashboard's Executive Summary for the full breakdown and recommendations.
