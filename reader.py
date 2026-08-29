# -*- coding: utf-8 -*-
"""
محوّل بيانات داشبورد المختصون
يقرأ ملف الإكسل "تقرير اعمال المختصون ... xlsx" ويولّد ملف data.js
الاستخدام:  python reader.py  [path-to-xlsx]
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
DEFAULT_SRC = r"C:\Users\GH\Desktop\تقرير اعمال المختصون يونيو.xlsx"


def find_source(arg=None):
    if arg and os.path.exists(arg):
        return arg
    if os.path.exists(DEFAULT_SRC):
        return DEFAULT_SRC
    hits = sorted(
        glob.glob(r"C:\Users\GH\Desktop\*.xlsx"),
        key=os.path.getmtime,
        reverse=True,
    )
    hits = [h for h in hits if "مختص" in h or "ملف" in h]
    if hits:
        return hits[0]
    raise FileNotFoundError("لم أجد ملف الإكسل. مرّر المسار كمعامل: python reader.py الملف.xlsx")


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
    """Iso الشهر/السنة الموجودين كرقم أو كتاريخ."""
    m = num(row[mcol]) if mcol is not None else 0
    y = num(row[ycol]) if ycol is not None else 0
    return int(m), int(y)


# ======================================================================
#  1) تقرير الإدارة  ->  بطاقة الإسناد / الإنتاجية / الفوترة الشهرية
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

# (label, first_month_row, last_month_row, total_row) — يُرمَّز بعد الكشف الديناميكي
ADM_ORDER = ["agg", "2023", "2024", "2025", "2026_قديم", "2026_جديد"]
YEAR_MAP = {"agg": "agg", "2023": 2023, "2024": 2024, "2025": 2025,
            "2026_قديم": 2026, "2026_جديد": 2026}


def find_adm_blocks(adm):
    """يكشف كتل السنوات ديناميكياً بدلاً من الصفوف الثابتة (يسمح بإضافة/حذف صفوف).
    كل كتلة تبدأ بسطر 'الشهر' تليه صفوف شهور وتنتهي بسطر 'الإجمالي'."""
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
    """إرجاع: punchcard[year_label][office][month] = [اسناد, انتاجيه, فوتره]"""
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


# ======================================================================
#  2) تقرير الفواتير  (المصروفات - المفوترة - ما تم صرفه)
# ======================================================================
def read_invoices(inv):
    """إرجاع هيكل الفواتير: تلخيص + تفاصيل شهرية."""
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

    # التفاصيل (تبدأ من الصف 9)
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

    # إزالة الفواتير المكررة تماماً (نفس كل الحقول بما فيها الأرقام التعريفية)
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
#  3) أوراق SAP و UDS  (إجماليات + حسب السنة / المكتب)
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
#  4) UNBILLED  (غير المفوتر)
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
#  5) تقرير تفصيلي  ->  الإسناد السنوي حسب القسم UDS/SAP
# ======================================================================
def read_type_split(tf):
    """كل سنة -> {UDS, SAP} من جدول تفاصيل الأعمال.""" 
    out = {}
    # عقد قديم: صف 4 فيه القيم منذ 2023 حتى 2026 (أعمدة UDS ثم SAP)
    base = tf.iloc[4]
    years = [2023, 2024, 2025, 2026]
    for k, y in enumerate(years):
        c = 1 + k * 2
        out[y] = {"UDS": num(base[c]), "SAP": num(base[c + 1])}
    # عقد جديد: صف 13 القيم تبدأ من 2026
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
#  تجميع
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

    # تصحيح الفواتير المكررة: خصم قيمتها من شهر/مكتب المطابق في البطاقة
    # (مرة واحدة لكل فاتورة فريدة، وليس لكل نسخة مكررة — حتى لا تُخصم مرتين)
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

    # ترتيب الصفوف في بيانات الفواتير/التفاصيل حسب الشهر ثم السنة
    invoices["detail"].sort(key=lambda d: (d["سنه"], d["شهر"]))

    data = {
        "meta": {
            "title": "ملخص أعمال المختصون",
            "source": os.path.basename(src),
            "generated_at": datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
            "years": ["agg", 2023, 2024, 2025, 2026],
            "offices": list(OFFICE_BLOCKS.keys()),
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

    # إضافة بصمة إصدار لملفات الأصول كي لا يحتفظ المتصفح بنسخة قديمة عند كل تحديث
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

    pc = data["punchcard"]["agg"]
    print("== التحقق من الأرقام ==")
    print("إجمالي الإسناد (agg):", data["punchcard"]["agg"]["مكتب خريص"]["_total"][0] +
          data["punchcard"]["agg"]["مكتب الشمال"]["_total"][0] +
          data["punchcard"]["agg"]["مكتب الجنوب"]["_total"][0] +
          data["punchcard"]["agg"]["طوارئ SAP"]["_total"][0])
    print("إجمالي الإنتاجية:", data["punchcard"]["agg"]["مكتب خريص"]["_total"][1] +
          data["punchcard"]["agg"]["مكتب الشمال"]["_total"][1] +
          data["punchcard"]["agg"]["مكتب الجنوب"]["_total"][1] +
          data["punchcard"]["agg"]["طوارئ SAP"]["_total"][1])
    print("إجمالي الفوترة:", data["punchcard"]["agg"]["مكتب خريص"]["_total"][2] +
          data["punchcard"]["agg"]["مكتب الشمال"]["_total"][2] +
          data["punchcard"]["agg"]["مكتب الجنوب"]["_total"][2] +
          data["punchcard"]["agg"]["طوارئ SAP"]["_total"][2])
    print("الفواتير شامل الضريبة:", data["invoices"]["foutra_shamel"])
    print("ما تم صرفه:", data["invoices"]["paid_total"], "| تحقق من التفاصيل:", data["invoices"]["_check"])
    print("SAP اسناد:", data["sap"]["totals"]["اسناد"], "مفوتر:", data["sap"]["totals"]["مفوتر"],
          "شامل:", data["sap"]["totals"]["shamel"])
    print("UDS اسناد:", data["uds"]["totals"]["اسناد"], "مفوتر بدون:", data["uds"]["totals"]["مفوتر"],
          "صرف:", data["uds"]["totals"]["صرف"], "متبقي صرف:", data["uds"]["totals"]["متبقي"])
    print("غير المفوتر:", data["unbilled"]["total"], "| مشاريع:", len(data["unbilled"]["projects"]))
    print("تقسيم UDS/SAP سنوياً:", data["type_split"])
    print("\nتم توليد:", out)


if __name__ == "__main__":
    main()