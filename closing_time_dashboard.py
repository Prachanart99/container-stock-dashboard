#!/usr/bin/env python3
"""
closing_time_dashboard.py
--------------------------
Regenerates the "Schedule & Closing Time" (N-Bound) dashboard from three
daily exports:

  SCHEDULE  : daily vessel schedule export (.xls) with columns
              Service, Vessel, Vessel Name, Op.Liner, Seq, Vyg, Bound,
              Vyg Bound, Wharf, POL, POD, Skip, No D, No L, USED,
              ETA Date, ETB Date, ETD Date
  INBOUND   : daily inbound (import) booking export (.xls)
  OUTBOUND  : daily outbound (export) booking export (.xls)

Both booking exports need at minimum: VSL, VOY, LWharf (outbound) /
DWharf (inbound), and the per-type container columns 12GP, 22GP, 42GP,
45GP, 12RE, 22RE, 42RE, 45RE, 22UT, 42UT, 22PC, 42PC, 45PC, 22TN, 42TN.

Only Bound == 'N' rows are kept (Northbound / outbound calls), and rows
with Skip == 'Y' are dropped. Rows are grouped into tabs by the first 3
letters of Wharf (BKK / LCH).

Closing time rule (fixed, matches the dashboard's own footnote):
  BKK tab: Dry = ETB - 12h, Reefer = ETB - 6h
  LCH tab: Dry = ETA - 24h, Reefer = ETA - 1h
A closing time in the past (relative to --now, default = current time in
Asia/Bangkok) renders as a red "Closed" pill; otherwise green "Open".

Inbound/Outbound (cntr) columns are a single container count (not TEU):
  Out Qty : sum of container-type columns from OUTBOUND where
            VSL == Vessel and VOY == Vyg Bound (exact, e.g. "2610N")
            and LWharf starts with the tab's wharf prefix.
  In Qty  : sum of container-type columns from INBOUND where
            VSL == Vessel and the numeric part of VOY == Vyg (bound
            letter stripped, since the inbound leg of a voyage carries
            the opposite direction suffix) and DWharf starts with the
            tab's wharf prefix.
No matching booking at all renders as "-" (na), not 0.

This script never invents the dashboard's HTML/CSS/JS shell — it patches
an existing dashboard export in place (--html), replacing only: the two
tab-count badges, both tables' <tbody> contents, the calendarEvents
object, and the reportNow timestamp used for open/closed + calendar
"today" logic. Point --html at any previous day's generated dashboard.

Usage
-----
    python3 closing_time_dashboard.py SCHEDULE.xls INBOUND.xls OUTBOUND.xls \\
        --html Closing_Time_NBound_prev.html --out Closing_Time_NBound_2026-09-18.html

Requires: xlrd (pip install xlrd --break-system-packages)
"""

import argparse
import datetime
import json
import re
from collections import defaultdict

import xlrd

BREAKDOWN_COLS = [
    "12GP", "22GP", "42GP", "45GP", "12RE", "22RE", "42RE", "45RE",
    "22UT", "42UT", "22PC", "42PC", "45PC", "22TN", "42TN",
]


def load_sheet(path):
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    headers = [sh.cell_value(0, c) for c in range(sh.ncols)]
    idx = {h: i for i, h in enumerate(headers)}
    rows = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(1, sh.nrows)]
    return idx, rows


def num(v):
    return v if isinstance(v, (int, float)) else 0


def parse_dt(s):
    return datetime.datetime.strptime(str(s).strip(), "%Y-%m-%d %H:%M")


def fmt_dt(dt):
    return dt.strftime("%d %b, %H:%M")


def numeric_voy(v):
    return re.sub(r"[A-Za-z]", "", str(v))


def build_booking_index(rows, idx, wharf_col, voy_key):
    """voy_key(row) -> the value to match Vyg/Vyg Bound against."""
    sums, counts = defaultdict(float), defaultdict(int)
    for r in rows:
        key = (r[idx["VSL"]], voy_key(r, idx), str(r[idx[wharf_col]])[:3])
        sums[key] += sum(num(r[idx[c]]) for c in BREAKDOWN_COLS)
        counts[key] += 1
    return sums, counts


def build_row(r, sidx, out_sums, out_counts, in_sums, in_counts, report_now):
    wpref = str(r[sidx["Wharf"]])[:3]
    vessel = r[sidx["Vessel"]]
    vyg = r[sidx["Vyg"]]
    voybound = r[sidx["Vyg Bound"]]
    eta = parse_dt(r[sidx["ETA Date"]])
    etb = parse_dt(r[sidx["ETB Date"]])
    etd = parse_dt(r[sidx["ETD Date"]])

    if wpref == "BKK":
        dry_close = etb - datetime.timedelta(hours=12)
        reefer_close = etb - datetime.timedelta(hours=6)
    else:
        dry_close = eta - datetime.timedelta(hours=24)
        reefer_close = eta - datetime.timedelta(hours=1)

    dry_status = "closed" if dry_close <= report_now else "open"
    reefer_status = "closed" if reefer_close <= report_now else "open"

    key_out = (vessel, voybound, wpref)
    key_in = (vessel, numeric_voy(vyg), wpref)
    outqty = int(round(out_sums[key_out])) if out_counts.get(key_out, 0) > 0 else None
    inqty = int(round(in_sums[key_in])) if in_counts.get(key_in, 0) > 0 else None

    return dict(
        wpref=wpref, service=r[sidx["Service"]], opliner=r[sidx["Op.Liner"]],
        vessel=vessel, vyg=vyg, vname=r[sidx["Vessel Name"]], voybound=voybound,
        eta=eta, etb=etb, etd=etd, dry_close=dry_close, reefer_close=reefer_close,
        dry_status=dry_status, reefer_status=reefer_status, inqty=inqty, outqty=outqty,
    )


def render_tbody(rows):
    out = []
    for i, row in enumerate(rows):
        cls = "band" if i % 2 == 1 else ""
        out.append(f'            <tr class="{cls}">')
        out.append(f'              <td class="mono">{row["service"]}</td>')
        out.append(f'              <td class="mono">{row["opliner"]}</td>')
        out.append(f'              <td class="mono">{row["vessel"]}</td>')
        out.append(f'              <td class="mono">{row["vyg"]}</td>')
        out.append(f'              <td class="vname">{row["vname"]}</td>')
        out.append(f'              <td class="mono voybound">{row["voybound"]}</td>')
        out.append(
            f'              <td class="closing dry {row["dry_status"]}">'
            f'<div class="dt">{fmt_dt(row["dry_close"])}</div>'
            f'<span class="pill {row["dry_status"]}" data-i18n="pill_{row["dry_status"]}"></span></td>'
        )
        out.append(
            f'              <td class="closing reefer {row["reefer_status"]}">'
            f'<div class="dt">{fmt_dt(row["reefer_close"])}</div>'
            f'<span class="pill {row["reefer_status"]}" data-i18n="pill_{row["reefer_status"]}"></span></td>'
        )
        out.append(f'              <td>{fmt_dt(row["eta"])}</td>')
        out.append(f'              <td>{fmt_dt(row["etb"])}</td>')
        out.append(f'              <td>{fmt_dt(row["etd"])}</td>')
        in_val = row["inqty"]
        out_val = row["outqty"]
        in_cls = "mono cntr" + (" na" if in_val is None else " ")
        out_cls = "mono cntr" + (" na" if out_val is None else " ")
        out.append(f'              <td class="{in_cls}">{in_val if in_val is not None else "—"}</td>')
        out.append(f'              <td class="{out_cls}">{out_val if out_val is not None else "—"}</td>')
        out.append("            </tr>")
    return "\n".join(out)


def build_calendar_events(rows, report_now):
    events = []
    for row in rows:
        status = "past" if row["etb"] < report_now else "upcoming"
        events.append({
            "date": row["etb"].strftime("%Y-%m-%d"),
            "type": "etb",
            "status": status,
            "vessel": row["vname"],
            "voyBound": row["voybound"],
            "time": fmt_dt(row["etb"]),
            "inbound": row["inqty"],
            "outbound": row["outqty"],
        })
    return events


def patch_html(html, bkk_rows, lch_rows, report_now):
    html = re.sub(
        r'(data-tab="bkk">BKK <span class="count">)\d+(</span>)',
        rf"\g<1>{len(bkk_rows)}\g<2>", html, count=1)
    html = re.sub(
        r'(data-tab="lch">LCH <span class="count">)\d+(</span>)',
        rf"\g<1>{len(lch_rows)}\g<2>", html, count=1)

    tbody_pattern = re.compile(r"(<tbody>\s*\n)(.*?)(\n\s*</tbody>)", re.DOTALL)
    matches = list(tbody_pattern.finditer(html))
    if len(matches) != 2:
        raise ValueError(f"expected 2 <tbody> blocks (BKK, LCH), found {len(matches)}")

    bkk_tbody = render_tbody(bkk_rows)
    lch_tbody = render_tbody(lch_rows)

    parts = []
    last_end = 0
    for m, new_body in zip(matches, (bkk_tbody, lch_tbody)):
        parts.append(html[last_end:m.start(2)])
        parts.append(new_body)
        last_end = m.end(2)
    parts.append(html[last_end:])
    html = "".join(parts)

    report_now_js = f'var reportNow = new Date("{report_now.strftime("%Y-%m-%dT%H:%M:%S")}");'
    html = re.sub(r'var reportNow = new Date\("[^"]*"\);', report_now_js, html, count=1)

    calendar_events = {
        "bkk": build_calendar_events(bkk_rows, report_now),
        "lch": build_calendar_events(lch_rows, report_now),
    }
    calendar_js = "var calendarEvents = " + json.dumps(calendar_events, ensure_ascii=False) + ";"
    html = re.sub(r"var calendarEvents = \{.*?\};", lambda m: calendar_js, html, count=1, flags=re.DOTALL)

    return html


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("schedule_file", help="Daily vessel schedule export (.xls)")
    ap.add_argument("inbound_file", help="Daily inbound (import) booking export (.xls)")
    ap.add_argument("outbound_file", help="Daily outbound (export) booking export (.xls)")
    ap.add_argument("--html", required=True, help="Existing dashboard HTML export to patch (shell reused as-is)")
    ap.add_argument("--out", required=True, help="Output path for the patched HTML")
    ap.add_argument("--now", help="Override report generation time, e.g. '2026-09-17 15:05:00' "
                                   "(default: current time, Asia/Bangkok UTC+7)")
    args = ap.parse_args()

    if args.now:
        report_now = datetime.datetime.strptime(args.now, "%Y-%m-%d %H:%M:%S")
    else:
        report_now = datetime.datetime.utcnow() + datetime.timedelta(hours=7)

    sidx, srows = load_sheet(args.schedule_file)
    nbound = [r for r in srows if r[sidx["Bound"]] == "N" and r[sidx["Skip"]] != "Y"]

    oidx, orows = load_sheet(args.outbound_file)
    iidx, irows = load_sheet(args.inbound_file)

    out_sums, out_counts = build_booking_index(orows, oidx, "LWharf", lambda r, idx: r[idx["VOY"]])
    in_sums, in_counts = build_booking_index(irows, iidx, "DWharf", lambda r, idx: numeric_voy(r[idx["VOY"]]))

    rows_by_wharf = defaultdict(list)
    for r in nbound:
        row = build_row(r, sidx, out_sums, out_counts, in_sums, in_counts, report_now)
        rows_by_wharf[row["wpref"]].append(row)

    bkk_rows = rows_by_wharf.get("BKK", [])
    lch_rows = rows_by_wharf.get("LCH", [])

    with open(args.html, encoding="utf-8") as f:
        html = f.read()
    patched = patch_html(html, bkk_rows, lch_rows, report_now)

    with open(args.out, "w", encoding="utf-8") as f:
        f.write(patched)

    print(f"BKK calls: {len(bkk_rows)}  LCH calls: {len(lch_rows)}")
    print(f"Report time: {report_now.isoformat()}")
    print(f"Written to: {args.out}")


if __name__ == "__main__":
    main()
