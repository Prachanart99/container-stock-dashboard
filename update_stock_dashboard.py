#!/usr/bin/env python3
"""
update_stock_dashboard.py
--------------------------
Reads the two daily container-stock files (one from SINOKOR, one from
HEUNG A) and either:
  1. prints the JavaScript `DATA` / `REEFER` blocks ready to paste into
     the "Container Stock Condition" dashboard's <script>, or
  2. patches those blocks directly into a copy of the dashboard's HTML
     file, if you pass --html.

Expected inputs
---------------
SINOKOR file  : has a sheet named "New Format"
                (blocks titled "SKR-THBKK" / "SKR-THLCH")
                optional sheet named "REEFER" for 45RE fleet-by-year detail
HEUNG A file  : has a sheet named "NEW FORMAT"
                (blocks titled "HAL (THBKK)" / "HAL (THLCH)")
                optional sheet named "RF SEASONAL" for 45RE fleet-by-year detail

Row labels are searched for by text (not fixed row numbers), so a blank
row or two between labels (as HEUNG A's sheet has) is handled fine.

IMPORTANT — about the Normal/Tight/Short/Surplus "condition" labels:
The daily per-carrier files above do NOT include a CONDITION row (only
the older combined stock_status_both.xlsx did). Per-type par levels
(balance thresholds), as given by ops on 18-Sep-2026, are hardcoded in
PAR_LEVELS below and drive the auto classification for 22GP, 42GP,
45GP, 22RE and 45RE. The remaining types (22UT, 42UT, 22PC, 42PC) have
no known par levels, so they still only get Short (balance <= 0)
auto-detected and default to Normal otherwise — printed as a warning
list for you to review/adjust by hand.

Usage
-----
    python3 update_stock_dashboard.py SINOKOR.xlsx HEUNGA.xlsx
    python3 update_stock_dashboard.py SINOKOR.xlsx HEUNGA.xlsx --html stock_status.html --out updated.html
    python3 update_stock_dashboard.py SINOKOR.xlsx HEUNGA.xlsx --date "16 Sep 2026"

Requires: openpyxl  (pip install openpyxl --break-system-packages)
"""

import argparse
import datetime
import re
import sys

import openpyxl

TYPES = ["22GP", "42GP", "45GP", "22RE", "45RE", "22UT", "42UT", "22PC", "42PC"]

# Par levels given by ops (18-Sep-2026), balance thresholds per type.
# Each entry is a list of (min_balance, label) checked high-to-low; the
# first threshold the balance meets or exceeds wins. Types not listed
# here have no known par levels (see classify()).
PAR_LEVELS = {
    "22GP": [(800, "Normal"), (500, "Tight"), (float("-inf"), "Short")],
    "42GP": [(2, "Normal"), (1, "Tight"), (float("-inf"), "Short")],
    "45GP": [(1000, "Surplus"), (700, "Normal"), (400, "Tight"), (float("-inf"), "Short")],
    "22RE": [(10, "Normal"), (float("-inf"), "Short")],
    "45RE": [(80, "Normal"), (20, "Tight"), (float("-inf"), "Short")],
}


def none0(v):
    return v if isinstance(v, (int, float)) else 0


def norm(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def find_row(ws, label, col, start_row, end_row):
    """Row number whose cell in `col` matches `label`, searched only within [start_row, end_row]."""
    target = norm(label)
    for r in range(start_row, end_row + 1):
        v = ws.cell(row=r, column=col).value
        if v is not None and norm(v) == target:
            return r
    raise ValueError(f"Row labeled {label!r} not found in column {col} between rows {start_row}-{end_row}")


def find_rows_starting(ws, prefix, col, start_row, end_row):
    """All row numbers whose cell in `col` starts with `prefix`, within [start_row, end_row], in order."""
    target = norm(prefix)
    out = []
    for r in range(start_row, end_row + 1):
        v = ws.cell(row=r, column=col).value
        if v is not None and norm(v).startswith(target):
            out.append(r)
    return out


def read_9(ws, row, col_start):
    return [none0(ws.cell(row=row, column=col_start + j).value) for j in range(9)]


def extract_sinokor_block(ws, title_row, next_title_row, col_start=2):
    """SINOKOR 'New Format' block: label col = col_start-1, values start at col_start.
    Searches only between title_row and next_title_row (exclusive) so BKK/LCH never collide."""
    label_col = col_start - 1
    end = next_title_row - 1
    stock_row = find_row(ws, "Current Stock", label_col, title_row, end)
    mt_row = find_row(ws, "IED(Estimated)", label_col, title_row, end)
    booking_rows = find_rows_starting(ws, "Booking (", label_col, title_row, end)
    if len(booking_rows) < 2:
        raise ValueError(f"Expected 2 'Booking (...)' rows between rows {title_row}-{end}, found {len(booking_rows)}")
    balance_row = find_row(ws, "Balance", label_col, title_row, end)
    return {
        "stock": read_9(ws, stock_row, col_start),
        "mtReturn": read_9(ws, mt_row, col_start),
        "bookThis": read_9(ws, booking_rows[0], col_start),
        "bookNext": read_9(ws, booking_rows[1], col_start),
        "balance": read_9(ws, balance_row, col_start),
    }


def extract_heunga_block(ws, title_row, next_title_row, col_start=3):
    label_col = col_start - 1
    end = next_title_row - 1
    stock_row = find_row(ws, "Current Stock", label_col, title_row, end)
    mt_row = find_row(ws, "INBOUND MT RETURN", label_col, title_row, end)
    bk_this_row = find_row(ws, "Booking(This week)", label_col, title_row, end)
    bk_next_row = find_row(ws, "Booking(Next week)", label_col, title_row, end)
    balance_row = find_row(ws, "Balance", label_col, title_row, end)
    return {
        "stock": read_9(ws, stock_row, col_start),
        "mtReturn": read_9(ws, mt_row, col_start),
        "bookThis": read_9(ws, bk_this_row, col_start),
        "bookNext": read_9(ws, bk_next_row, col_start),
        "balance": read_9(ws, balance_row, col_start),
    }


def classify(type_name, balance):
    """Classify a type's balance using ops' PAR_LEVELS when known; otherwise
    fall back to the only label this script can set with confidence:
    Short = balance <= 0, everything else defaults to Normal."""
    levels = PAR_LEVELS.get(type_name)
    if levels is None:
        return "Short" if balance <= 0 else "Normal"
    for threshold, label in levels:
        if balance >= threshold:
            return label
    return "Short"


def extract_summary(sinokor_path, heunga_path):
    wb1 = openpyxl.load_workbook(sinokor_path, data_only=True)
    ws1 = wb1["New Format"]
    snko_bkk_title = find_row(ws1, "SKR-THBKK", 1, 1, 10)
    snko_lch_title = find_row(ws1, "SKR-THLCH", 1, snko_bkk_title + 1, snko_bkk_title + 30)
    snko_bkk = extract_sinokor_block(ws1, snko_bkk_title, snko_lch_title)
    snko_lch = extract_sinokor_block(ws1, snko_lch_title, snko_lch_title + 30)

    wb2 = openpyxl.load_workbook(heunga_path, data_only=True)
    ws2 = wb2["NEW FORMAT"]
    hal_bkk_title = find_row(ws2, "HAL (THBKK)", 2, 1, 10)
    hal_lch_title = find_row(ws2, "HAL (THLCH)", 2, hal_bkk_title + 1, hal_bkk_title + 30)
    hal_bkk = extract_heunga_block(ws2, hal_bkk_title, hal_lch_title)
    hal_lch = extract_heunga_block(ws2, hal_lch_title, hal_lch_title + 30)

    review_needed = []
    for name, block in (("SNKO_BKK", snko_bkk), ("HAL_BKK", hal_bkk),
                         ("SNKO_LCH", snko_lch), ("HAL_LCH", hal_lch)):
        block["condition"] = [classify(t, b) for t, b in zip(TYPES, block["balance"])]
        for t, bal, cond in zip(TYPES, block["balance"], block["condition"]):
            if t not in PAR_LEVELS and cond == "Normal" and bal > 0:
                review_needed.append((name, t, bal))

    return {
        "SNKO_BKK": {"carrier": "SINOKOR", "port": "BKK", "color": "sinokor", **snko_bkk},
        "HAL_BKK": {"carrier": "HEUNG A", "port": "BKK", "color": "heunga", **hal_bkk},
        "SNKO_LCH": {"carrier": "SINOKOR", "port": "LCH", "color": "sinokor", **snko_lch},
        "HAL_LCH": {"carrier": "HEUNG A", "port": "LCH", "color": "heunga", **hal_lch},
    }, review_needed


def extract_reefer_sheet(ws, year_col, val_cols, total_row_label="TOTAL", search_range=20):
    rows = []
    total = None
    for r in range(1, search_range + 1):
        cell = ws.cell(row=r, column=year_col).value
        if cell is None:
            continue
        if str(cell).strip().upper() == total_row_label:
            total = {
                "carrier": none0(ws.cell(row=r, column=val_cols[0]).value),
                "daikin": none0(ws.cell(row=r, column=val_cols[1]).value),
                "thermoking": none0(ws.cell(row=r, column=val_cols[2]).value),
            }
            continue
        try:
            year = int(cell)
        except (ValueError, TypeError):
            continue
        carrier = none0(ws.cell(row=r, column=val_cols[0]).value)
        daikin = none0(ws.cell(row=r, column=val_cols[1]).value)
        thermoking = none0(ws.cell(row=r, column=val_cols[2]).value)
        if carrier or daikin or thermoking:
            rows.append({"year": year, "carrier": carrier, "daikin": daikin, "thermoking": thermoking})
    rows.sort(key=lambda x: -x["year"])
    if total is None:
        total = {
            "carrier": sum(r["carrier"] for r in rows),
            "daikin": sum(r["daikin"] for r in rows),
            "thermoking": sum(r["thermoking"] for r in rows),
        }
    return {"rows": rows, "total": total}


def extract_reefer(sinokor_path, heunga_path):
    result = {}
    try:
        wb1 = openpyxl.load_workbook(sinokor_path, data_only=True)
        ws1 = wb1["REEFER"]
        result["SNKO_BKK"] = {"carrier": "SINOKOR", "color": "sinokor",
                               **extract_reefer_sheet(ws1, 1, (2, 3, 4))}
        result["SNKO_LCH"] = {"carrier": "SINOKOR", "color": "sinokor",
                               **extract_reefer_sheet(ws1, 6, (7, 8, 9))}
    except KeyError:
        print("Note: no 'REEFER' sheet found in SINOKOR file — skipping reefer fleet detail for SINOKOR.",
              file=sys.stderr)

    try:
        wb2 = openpyxl.load_workbook(heunga_path, data_only=True)
        ws2 = wb2["RF SEASONAL"]
        result["HAL_BKK"] = {"carrier": "HEUNG A", "color": "heunga",
                              **extract_reefer_sheet(ws2, 1, (2, 3, 4))}
        result["HAL_LCH"] = {"carrier": "HEUNG A", "color": "heunga",
                              **extract_reefer_sheet(ws2, 6, (7, 8, 9))}
    except KeyError:
        print("Note: no 'RF SEASONAL' sheet found in HEUNG A file — skipping reefer fleet detail for HEUNG A.",
              file=sys.stderr)

    return result


def _format_date_value(v):
    if isinstance(v, datetime.datetime):
        return v.strftime("%d %b %Y").lstrip("0")
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(s, fmt).strftime("%d %b %Y").lstrip("0")
        except ValueError:
            continue
    return s  # fall back to whatever text was there


def find_report_date(*paths):
    """Look for a 'DATE:' label near the top of the Daily sheet (trying each path in turn)
    and read the value beside it, formatted as e.g. '16 Sep 2026'."""
    for path in paths:
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb["Daily"] if "Daily" in wb.sheetnames else wb[wb.sheetnames[0]]
            for row in ws.iter_rows(min_row=1, max_row=3):
                for i, cell in enumerate(row):
                    if cell.value and "DATE" in str(cell.value).upper():
                        for other in row[i + 1:]:
                            if other.value:
                                return _format_date_value(other.value)
        except Exception:
            continue
    return None


def js_array(vals):
    return "[" + ",".join(str(v) for v in vals) + "]"


def render_data_js(summary):
    lines = ["const DATA = {"]
    for key in ("SNKO_BKK", "HAL_BKK", "SNKO_LCH", "HAL_LCH"):
        b = summary[key]
        lines.append(f"  {key}: {{ carrier:'{b['carrier']}', port:'{b['port']}', color:'{b['color']}',")
        lines.append(f"    stock:{js_array(b['stock'])},")
        lines.append(f"    mtReturn:{js_array(b['mtReturn'])},")
        lines.append(f"    bookThis:{js_array(b['bookThis'])}, bookNext:{js_array(b['bookNext'])},")
        lines.append(f"    balance:{js_array(b['balance'])},")
        cond = ",".join(f"'{c}'" for c in b["condition"])
        lines.append(f"    condition:[{cond}] }},")
    lines.append("};")
    return "\n".join(lines)


def render_reefer_js(reefer):
    if not reefer:
        return "const REEFER = {\n};"
    lines = ["const REEFER = {"]
    for key, b in reefer.items():
        lines.append(f"  {key}: {{ carrier:'{b['carrier']}', color:'{b['color']}', rows:[")
        for r in b["rows"]:
            lines.append(f"    {{ year:{r['year']}, carrier:{r['carrier']}, daikin:{r['daikin']}, thermoking:{r['thermoking']} }},")
        t = b["total"]
        lines.append(f"  ], total:{{ carrier:{t['carrier']}, daikin:{t['daikin']}, thermoking:{t['thermoking']} }} }},")
    lines.append("};")
    return "\n".join(lines)


def patch_html(html_text, data_js, reefer_js, date_str=None):
    html_text = re.sub(r"const DATA = \{[\s\S]*?\n\};", data_js, html_text, count=1)
    if reefer_js.strip() != "const REEFER = {\n};":
        html_text = re.sub(r"const REEFER = \{[\s\S]*?\n\};", reefer_js, html_text, count=1)
    if date_str:
        html_text = re.sub(r"(date:\s*')[^']*(')", rf"\g<1>{date_str}\g<2>", html_text, count=1)
        html_text = re.sub(
            r'(<span class="date mono" id="t-date">)[^<]*(</span>)',
            rf"\g<1>{date_str}\g<2>",
            html_text,
            count=1,
        )
    return html_text


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("sinokor_file", help="SINOKOR daily stock .xlsx (has a 'New Format' sheet)")
    ap.add_argument("heunga_file", help="HEUNG A daily stock .xlsx (has a 'NEW FORMAT' sheet)")
    ap.add_argument("--html", help="Existing dashboard HTML file to patch in place (or write to --out)")
    ap.add_argument("--out", help="Output path for the patched HTML (defaults to overwriting --html)")
    ap.add_argument("--date", help="Override the report date label, e.g. '16 Sep 2026'")
    args = ap.parse_args()

    summary, review_needed = extract_summary(args.sinokor_file, args.heunga_file)
    reefer = extract_reefer(args.sinokor_file, args.heunga_file)
    date_str = args.date or find_report_date(args.sinokor_file, args.heunga_file)

    data_js = render_data_js(summary)
    reefer_js = render_reefer_js(reefer)

    if args.html:
        with open(args.html, "r", encoding="utf-8") as f:
            html_text = f.read()
        patched = patch_html(html_text, data_js, reefer_js, date_str)
        out_path = args.out or args.html
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(patched)
        print(f"Patched dashboard written to: {out_path}")
        if date_str:
            print(f"Report date set to: {date_str}")
    else:
        if date_str:
            print(f"// Report date detected: {date_str}\n")
        print(data_js)
        print()
        print(reefer_js)

    if review_needed:
        print("\n--- REVIEW NEEDED: these have balance > 0 and were defaulted to 'Normal' ---", file=sys.stderr)
        print("(the source file has no CONDITION row, so Tight/Surplus can't be auto-detected —", file=sys.stderr)
        print(" edit the condition:[...] array by hand for any of these that should be Tight/Surplus)", file=sys.stderr)
        for name, t, bal in review_needed:
            print(f"  {name} {t}: balance={bal}", file=sys.stderr)


if __name__ == "__main__":
    main()
