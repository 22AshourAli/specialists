# -*- coding: utf-8 -*-
"""
Mokhtasoon dashboard data converter.
Reads the monthly Excel file "تقرير اعمال المختصون ... xlsx" and generates data.js.
Usage: python reader.py [path-to-xlsx]
Tip: the latest month's file on the Desktop is auto-detected when no path is given.
"""
import sys
import os
import re
import json
import datetime
import glob
from datetime import datetime as _dt
from datetime import timedelta

import pandas as pd
import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
FILE_PREFIX = "تقرير اعمال المختصون"
NOTES = ["UNBILLED", "تقرير الاداره", "تقرير الفواتير", "SAP", "UDS", "تقرير تفصيلي"]
NAME_FIX = {"الدرعيه": "الدرعية"}  # sheet name -> clean Arabic office name


def _desktop_candidates():
    """Possible Desktop locations; the Windows profile can vary between environments."""
    candidates = []
    for var in ("USERPROFILE", "HOME"):
        base = os.environ.get(var)
        if base:
            candidates.append(os.path.join(base, "Desktop"))
    candidates.append(os.path.join(os.path.expanduser("~"), "Desktop"))
    for home in (r"C:\Users\GH", r"C:\Users\Ashour"):
        if os.path.isdir(home):
            candidates.append(os.path.join(home, "Desktop"))
    return candidates


def find_source(arg=None):
    """Resolve the Excel source: an explicit path, or the newest monthly file on the Desktop."""
    if arg and os.path.exists(arg):
        return arg
    for desk in _desktop_candidates():
        hits = sorted(
            glob.glob(os.path.join(desk, FILE_PREFIX + "*.xlsx")),
            key=os.path.getmtime,
            reverse=True,
        )
        if hits:
            return hits[0]
    raise FileNotFoundError("No dashboard file found on the Desktop. Pass a path: python reader.py file.xlsx")


def clean(v):
    if v is None:
        return None
    if isinstance(v, (int, float, np.integer, np.floating)):
        f = float(v)
        if np.isnan(f):
            return None
        return round(f, 4)
    s = str(v).strip()
    return s if s else None


def num(v):
    if v is None:
        return 0.0
    try:
        f = float(v)
        return f if not np.isnan(f) else 0.0
    except (TypeError, ValueError):
        return 0.0


def parse_month_year(row, mcol, ycol):
    """Extract the month/year pair stored as number or date in the given cells."""
    m = num(row[mcol]) if mcol is not None else 0
    y = num(row[ycol]) if ycol is not None else 0
    return int(m), int(y)


# ======================================================================
#  1) Admin report  ->  monthly esnad / entajia / foutra punchcard
# ======================================================================
OFFICE_BLOCKS = {
    "مكتب خريص": [("اسناد", 1), ("انتاجيه", 2), ("فوتره", 3)],
    "مكتب الشمال": [("اسناد", 4), ("انتاجيه", 5), ("فوتره", 6)],
    "مكتب الجنوب": [("اسناد", 7), ("انتاجيه", 8), ("فوتره", 9)],
    "طوارئ SAP": [
        ("اسناد", 10), ("اسناد", 13), ("اسناد", 16),
        ("انتاجيه", 11), ("انتاجيه", 14), ("انتاجيه", 17),
        ("فوتره", 12), ("فوتره", 15), ("فوتره", 18),
    ],
}
METRIC_KEYS = ["اسناد", "انتاجيه", "فوتره"]

# (label, first_month_row, last_month_row, total_row) — encoded after dynamic detection
ADM_ORDER = ["agg", "2023", "2024", "2025", "2026_قديم", "2026_جديد"]
YEAR_MAP = {"agg": "agg", "2023": 2023, "2024": 2024, "2025": 2025,
            "2026_قديم": 2026, "2026_جديد": 2026}


def find_adm_blocks(adm):
    """Detect the year blocks dynamically instead of fixed rows (allows added/removed rows).
    Each block starts with a 'الشهر' line, holds 'month' rows and ends with 'الإجمالي'."""
    blocks = []
    nrows = len(adm)

    def cell(row, col):
        v = adm.iloc[row, col]
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except Exception:
            pass
        return str(v).strip()

    for r in range(nrows):
        if cell(r, 0) != "الشهر":
            continue
        j = r + 2
        last_m = None
        tot = None
        while j < nrows:
            s = cell(j, 0)
            if s == "الإجمالي":
                tot = j
                break
            if s == "":
                j += 1
                continue
            try:
                n = int(round(float(s)))
                ok = 1 <= n <= 12 and re.match(r"^\d+(\.\d+)?$", s) is not None
            except ValueError:
                ok = False
            if ok:
                last_m = j
            else:
                break
            j += 1
        if last_m is not None and tot is not None:
            blocks.append((r + 2, last_m, tot))
    return blocks


def read_punchcard(adm):
    """Return punchcard[year_label][office][month] = [esnad, entajia, foutra]."""
    blocks = find_adm_blocks(adm)
    punch = {}
    for bi, (r0, r1, rtot) in enumerate(blocks):
        if bi >= len(ADM_ORDER):
            continue
        ylab = YEAR_MAP[ADM_ORDER[bi]]
        punch.setdefault(ylab, {})
        for off, spec in OFFICE_BLOCKS.items():
            punch[ylab].setdefault(off, {})
        for r in range(r0, r1 + 1):
            m = int(num(adm.iloc[r, 0]))
            if m < 1 or m > 12:
                continue
            for off, spec in OFFICE_BLOCKS.items():
                vals = {}
                for metric, col in spec:
                    vals[metric] = vals.get(metric, 0.0) + num(adm.iloc[r, col])
                prev = punch[ylab][off].get(m, [0.0, 0.0, 0.0])
                punch[ylab][off][m] = [
                    prev[0] + vals["اسناد"],
                    prev[1] + vals["انتاجيه"],
                    prev[2] + vals["فوتره"],
                ]
        for off, spec in OFFICE_BLOCKS.items():
            vals = {}
            for metric, col in spec:
                vals[metric] = vals.get(metric, 0.0) + num(adm.iloc[rtot, col])
            tot = punch[ylab][off].get("_total", [0.0, 0.0, 0.0])
            punch[ylab][off]["_total"] = [
                tot[0] + vals["اسناد"],
                tot[1] + vals["انتاجيه"],
                tot[2] + vals["فوتره"],
            ]
    return punch


def read_office_sheet(df):
    """Read annual assignment totals from a standalone office sheet ("خريص", "الشرق", ...).
    Layout: blocks titled "تفاصيل أوامر العمل المسنده ... عقد قديم/جديد" followed by
    السنه / القسم / القيمه / الإجمالي rows. Old contract: yearly UDS columns with a
    single SAP total; new contract: alternating UDS/SAP pairs per year."""
    nrows, ncols = df.shape

    def cell(r, c):
        if r < 0 or c < 0 or r >= nrows or c >= ncols:
            return ""
        v = df.iat[r, c]
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except Exception:
            pass
        return str(v).strip()

    blocks = []
    for r in range(nrows):
        t = cell(r, 0)
        if not t.startswith("تفاصيل أوامر العمل المسنده"):
            continue
        hdr, val = r + 1, r + 3
        if cell(hdr, 0) != "السنه" or cell(val, 0) != "القيمه":
            continue
        blocks.append({"old": "عقد قديم" in t, "hdr": hdr, "val": val})

    agg = 0.0
    yearly = {}
    for b in blocks:
        hdr, val = b["hdr"], b["val"]
        ycols = {}
        for c in range(1, ncols):
            if re.match(r"^20\d\d$", cell(hdr, c)):
                ycols[int(cell(hdr, c))] = c
        agg += num(df.iat[val, 9]) + num(df.iat[val, 10])
        for y, c in ycols.items():
            v = num(df.iat[val, c])
            if not b["old"] and c + 1 < ncols:
                v += num(df.iat[val, c + 1])
            yearly[y] = yearly.get(y, 0.0) + v
    # Note: old-contract SAP is only available as a total (already in `agg`);
    # per-year entries carry UDS (SAP is zero in the current files).
    return {
        "agg": round(agg, 4),
        "years": {k: round(v, 4) for k, v in yearly.items()},
    }


def office_extra_sheets(xl):
    """Map standalone office sheets (outside the standard set) to their yearly totals."""
    excluded = {n.strip() for n in NOTES}
    out = {}
    for sh in xl.sheet_names:
        st = str(sh).strip()
        if st in excluded or not st:
            continue
        label = "مكتب " + NAME_FIX.get(st, st)
        if label in OFFICE_BLOCKS:
            continue
        out[label] = read_office_sheet(xl.parse(sh, header=None))
    return out


# ======================================================================
#  2) Invoices report  (billed - invoiced - paid)
# ======================================================================
def read_invoices(inv):
    """Return the invoice structure: yearly summary + monthly detail."""
    inv_years = [2023, 2024, 2025, 2026]
    cols = {2023: 4, 2024: 7, 2025: 10, 2026: 13}
    summary = {}
    for y in inv_years:
        c = cols[y]
        summary[str(y)] = {
            "جزئية": round(num(inv.iloc[3, c]), 4),
            "نهائية": round(num(inv.iloc[4, c]), 4),
            "مباشر": round(num(inv.iloc[5, c]), 4),
            "إجمالي": round(num(inv.iloc[6, c]), 4),
        }

    def has_amt(i):
        return any(num(inv.iloc[i, c]) != 0 for c in range(11, 18))

    def fallback_yx(i, last):
        v = inv.iloc[i, 10]
        if v is not None and not pd.isna(v):
            if isinstance(v, _dt):
                return (v.month, v.year)
            if isinstance(v, (int, float)) and v > 1000:
                d = _dt(1899, 12, 30) + timedelta(days=v)
                return (d.month, d.year)
            if isinstance(v, str):
                m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.]\d{1,2}", v.strip())
                if m:
                    return (int(m.group(2)), int(m.group(1)))
                m = re.match(r"^(\d{1,2})[-/.]\d{1,2}[-/.](\d{4})", v.strip())
                if m:
                    return (int(m.group(1)), int(m.group(2)))
        return last

    # Detail starts at row 9
    detail = []
    max_year = 2026
    last = None
    for i in range(9, len(inv)):
        if pd.isna(inv.iloc[i, 0]) and pd.isna(inv.iloc[i, 2]):
            continue
        if not has_amt(i):
            if not str(inv.iloc[i, 0]).strip().isdigit():
                continue
            if num(inv.iloc[i, 9]) == 0:
                continue
        m, y = parse_month_year(inv.iloc[i], 8, 9)
        if y == 0 or m < 1 or m > 12:
            fy = fallback_yx(i, last)
            if fy:
                m, y = fy
            else:
                m, y = (last[0] if last else 12, max_year)
        last = (m, y)
        max_year = max(max_year, y)
        row = {
            "عقد": clean(inv.iloc[i, 1]),
            "قسم": clean(inv.iloc[i, 2]),
            "شهر": m,
            "سنه": y,
            "قيمة": round(num(inv.iloc[i, 11]), 4),
            "ضريبة": round(num(inv.iloc[i, 12]), 4),
            "shamel": round(num(inv.iloc[i, 13]), 4),
            "غرامات": round(num(inv.iloc[i, 14]), 4),
            "سلامه": round(num(inv.iloc[i, 15]), 4),
            "safy": round(num(inv.iloc[i, 16]), 4),
            "paid": round(num(inv.iloc[i, 17]), 4),
            "مكتب": clean(inv.iloc[i, 18]),
            "_sig": (
                clean(inv.iloc[i, 1]), clean(inv.iloc[i, 2]),
                clean(inv.iloc[i, 3]), clean(inv.iloc[i, 4]),
                clean(inv.iloc[i, 5]), clean(inv.iloc[i, 6]), clean(inv.iloc[i, 7]),
                m, y, str(inv.iloc[i, 10]),
                num(inv.iloc[i, 11]), num(inv.iloc[i, 12]), num(inv.iloc[i, 13]),
                num(inv.iloc[i, 14]), num(inv.iloc[i, 15]), num(inv.iloc[i, 16]),
                num(inv.iloc[i, 17]), clean(inv.iloc[i, 18]),
            ),
        }
        detail.append(row)

    # Drop fully-duplicated rows (same every field including identifying values)
    dropped = []
    seen = set()
    unique = []
    for row in detail:
        if row["_sig"] in seen:
            dropped.append(row)
            continue
        seen.add(row["_sig"])
        row.pop("_sig")
        unique.append(row)
    detail = unique

    foutra_bt = round(sum(r["قيمة"] for r in detail), 4)
    foutra_darib = round(sum(r["ضريبة"] for r in detail), 4)
    foutra_shamel = round(sum(r["shamel"] for r in detail), 4)
    paid_total = round(sum(r["paid"] for r in detail), 4)

    return {
        "by_year": summary,
        "paid_total": paid_total,
        "foutra_bt": foutra_bt,
        "foutra_darib": foutra_darib,
        "foutra_shamel": foutra_shamel,
        "detail": detail,
        "deduped_drops": dropped,
        "_check": round(sum(d["paid"] for d in detail), 4),
    }


# ======================================================================
#  3) SAP and UDS sheets  (totals + per year / office)
# ======================================================================
def read_sap(sap):
    rows = []
    for i in range(1, len(sap)):
        if not str(sap.iloc[i, 0]).strip().isdigit():
            continue
        m, y = parse_month_year(sap.iloc[i], 7, 8)
        if y == 0:
            continue
        rows.append({
            "مكتب": clean(sap.iloc[i, 1]),
            "عقد": clean(sap.iloc[i, 2]),
            "انوع": clean(sap.iloc[i, 5]),
            "موقف": clean(sap.iloc[i, 6]),
            "شهر": m, "سنه": y,
            "اسناد": num(sap.iloc[i, 17]),
            "مفوتر": num(sap.iloc[i, 18]),
            "ضريبه": num(sap.iloc[i, 19]),
            "shamel": num(sap.iloc[i, 20]),
            "صافي": num(sap.iloc[i, 23]),
        })
    totals = {
        "اسناد": round(sum(r["اسناد"] for r in rows), 4),
        "مفوتر": round(sum(r["مفوتر"] for r in rows), 4),
        "ضريبه": round(sum(r["ضريبه"] for r in rows), 4),
        "shamel": round(sum(r["shamel"] for r in rows), 4),
        "صافي": round(sum(r["صافي"] for r in rows), 4),
    }
    return {"totals": totals, "detail": rows}


def read_uds(uds):
    rows = []
    for i in range(1, len(uds)):
        if not str(uds.iloc[i, 0]).strip().isdigit():
            continue
        m, y = parse_month_year(uds.iloc[i], 27, 28)
        if y == 0:
            m, y = parse_month_year(uds.iloc[i], 13, 14)
        rows.append({
            "مكتب": clean(uds.iloc[i, 1]),
            "اسناد": num(uds.iloc[i, 24]),
            "مفوتر": num(uds.iloc[i, 35]),
            "shamel": num(uds.iloc[i, 31]),
            "صرف": num(uds.iloc[i, 36]),
            "متبقي": num(uds.iloc[i, 37]),
            "شهر": m, "سنه": y,
            "افادة_فوتره": clean(uds.iloc[i, 17]),
            "افادة_تنفيذ": clean(uds.iloc[i, 15]),
            "موقف": clean(uds.iloc[i, 18]),
        })
    totals = {
        "اسناد": round(sum(r["اسناد"] for r in rows), 4),
        "مفوتر": round(sum(r["مفوتر"] for r in rows), 4),
        "shamel": round(sum(r["shamel"] for r in rows), 4),
        "صرف": round(sum(r["صرف"] for r in rows), 4),
        "متبقي": round(sum(r["متبقي"] for r in rows), 4),
    }
    return {"totals": totals, "detail": rows}


# ======================================================================
#  4) UNBILLED  (unbilled revenue)
# ======================================================================
def read_unbilled(unb):
    projects = []
    for i in range(3, len(unb)):
        name = unb.iloc[i, 1]
        v = unb.iloc[i, 8]
        if pd.isna(name):
            continue
        projects.append({
            "name": clean(name),
            "مبدئيه": round(num(unb.iloc[i, 4]), 4),
            "نهائيه": round(num(unb.iloc[i, 5]), 4),
            "رصيد_مفوتر": round(num(v), 4),
        })
    total = round(sum(p["رصيد_مفوتر"] for p in projects), 4)
    return {"total": total, "projects": projects}


# ======================================================================
#  5) Type split  ->  annual esnad by section UDS/SAP
# ======================================================================
def read_type_split(tf):
    """Each year -> {UDS, SAP} read from the work-details table."""
    out = {}
    # Old contract: row 4 holds values for 2023..2026 (UDS then SAP columns)
    base = tf.iloc[4]
    years = [2023, 2024, 2025, 2026]
    for k, y in enumerate(years):
        c = 1 + k * 2
        out[y] = {"UDS": num(base[c]), "SAP": num(base[c + 1])}
    # New contract: row 13, values start at 2026
    if len(tf) > 13:
        base2 = tf.iloc[13]
        if num(base2[1]) or num(base2[2]):
            out[2026]["UDS"] = out[2026]["UDS"] + num(base2[1])
            out[2026]["SAP"] = out[2026]["SAP"] + num(base2[2])
    for y in out:
        out[y]["UDS"] = round(out[y]["UDS"], 4)
        out[y]["SAP"] = round(out[y]["SAP"], 4)
    return out


# ======================================================================
#  Assembly
# ======================================================================
def build(src):
    xl = pd.ExcelFile(src)
    adm = xl.parse("تقرير الاداره", header=None)
    inv = xl.parse("تقرير الفواتير", header=None)
    sap = xl.parse("SAP", header=None)
    uds = xl.parse("UDS", header=None)
    unb = xl.parse("UNBILLED", header=None)
    tf = xl.parse("تقرير تفصيلي ", header=None)

    punch = read_punchcard(adm)
    invoices = read_invoices(inv)
    sap = read_sap(sap)
    uds = read_uds(uds)
    unbilled = read_unbilled(unb)
    type_split = read_type_split(tf)

    # Extra offices (الشرق / الدرعية) only carry annual assignment totals - no monthly data.
    extra = office_extra_sheets(xl)
    for label, tot in extra.items():
        for yk in ["agg", 2023, 2024, 2025, 2026]:
            v = tot["agg"] if yk == "agg" else tot["years"].get(yk, 0.0)
            punch.setdefault(yk, {}).setdefault(label, {})["_total"] = [v, 0.0, 0.0]

    # Compensate duplicated invoices: subtract their value from the matching month/office
    # in the punchcard (once per unique invoice, not per duplicate copy).
    seen_d = set()
    for drop in invoices["deduped_drops"]:
        sig = drop.get("_sig")
        if sig is not None and sig in seen_d:
            continue
        if sig is not None:
            seen_d.add(sig)
        off = drop["مكتب"] or "طوارئ SAP"
        for yb in (drop["سنه"], "agg"):
            if yb in punch and off in punch[yb]:
                if drop["شهر"] in punch[yb][off]:
                    punch[yb][off][drop["شهر"]][2] -= drop["قيمة"]
                if "_total" in punch[yb][off]:
                    punch[yb][off]["_total"][2] -= drop["قيمة"]

    # Sort invoice/detail rows by year then month
    invoices["detail"].sort(key=lambda d: (d["سنه"], d["شهر"]))

    data = {
        "meta": {
            "title": "ملخص أعمال المختصون",
            "source": os.path.basename(src),
            "generated_at": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "years": ["agg", 2023, 2024, 2025, 2026],
            "offices": list(OFFICE_BLOCKS.keys()) + list(extra.keys()),
            "months": list(range(1, 13)),
            "counts": {
                "invoices": len(invoices["detail"]),
                "unbilled": len(unbilled["projects"]),
                "dups": len([d for d in invoices["deduped_drops"] if d.get("paid") or d.get("قيمة")]),
            },
        },
        "punchcard": punch,
        "invoices": invoices,
        "sap": sap,
        "uds": uds,
        "unbilled": unbilled,
        "type_split": type_split,
    }
    return data


def main():
    src = find_source(sys.argv[1] if len(sys.argv) > 1 else None)
    data = build(src)
    out = os.path.join(HERE, "data.js")
    js = "window.DASHBOARD_DATA = " + json.dumps(data, ensure_ascii=False, indent=1) + ";\n"
    with open(out, "w", encoding="utf-8") as f:
        f.write(js)

    # Stamp asset versions in index.html so browsers don't keep a stale cached copy
    stamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    idx = os.path.join(HERE, "index.html")
    with open(idx, "r", encoding="utf-8") as f:
        html = f.read()
    html = re.sub(
        r'((?:src|href)=")(assets/(?:style\.css|app\.js)|data\.js)(\?v=\d+)?(")',
        r"\1\2?v=" + stamp + r"\4",
        html,
    )
    with open(idx, "w", encoding="utf-8") as f:
        f.write(html)

    agg = data["punchcard"]["agg"]
    offices = data["meta"]["offices"]
    esn = sum(agg[o]["_total"][0] for o in offices)
    ent = sum(agg[o]["_total"][1] for o in offices)
    fou = sum(agg[o]["_total"][2] for o in offices)
    print("== number verification ==")
    print("total esnad (agg):", esn)
    print("total entajia (agg):", ent)
    print("total foutra (agg):", fou)
    print("offices:", offices)
    for o in offices:
        t = agg[o]["_total"]
        print("  -", o, "-> esnad", t[0], "| entajia", t[1], "| foutra", t[2])
    print("invoices incl. VAT:", data["invoices"]["foutra_shamel"])
    print("paid:", data["invoices"]["paid_total"], "| check:", data["invoices"]["_check"])
    print("SAP esnad:", data["sap"]["totals"]["اسناد"], "billed:", data["sap"]["totals"]["مفوتر"],
          "incl.:", data["sap"]["totals"]["shamel"])
    print("UDS esnad:", data["uds"]["totals"]["اسناد"], "billed w/o:", data["uds"]["totals"]["مفوتر"],
          "paid:", data["uds"]["totals"]["صرف"], "remaining:", data["uds"]["totals"]["متبقي"])
    print("unbilled:", data["unbilled"]["total"], "| projects:", len(data["unbilled"]["projects"]))
    print("type split by year:", data["type_split"])
    print("\nGenerated:", out)


if __name__ == "__main__":
    main()