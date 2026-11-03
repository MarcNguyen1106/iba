"""Verify BUSM1042_A3_Group_Clean_Data.xlsx and BUSM1042_A3_Data_Quality_Notes.docx.

Usage:
  python3 -I verify_group.py recalc <built.xlsx> <recalc_dir>          # LibreOffice: soffice --headless --calc --convert-to xlsx
  python3 -I verify_group.py check  <source.xlsx> <built.xlsx> <recalc.xlsx> <notes.docx> <notes.pdf> <report.md>
  python3 -I verify_group.py render <notes.docx> <out_dir>             # Word -> PDF -> PNG pages for the visual check

Method A = pandas on the ORIGINAL file; Method B = a plain Python row loop over the ORIGINAL file (openpyxl read-only,
no pandas); Method C = the workbook's own Excel formulas after LibreOffice recalculation, read with data_only=True.
Counts must agree exactly and rates within 1e-12; all three must equal the expected values in the brief.
"""
import os
import re
import subprocess
import sys
import time
import zipfile
from collections import Counter, defaultdict
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal

N = 119390
LAST = N + 1
OK = "✔ OK"
LEAD = ["0–7", "8–30", "31–90", "91–180", "181–365", "Over 365"]
NIGHTS = ["0", "1", "2–3", "4–7", "8+"]
ADR = ["0 or less", "0.01–50", "50.01–100", "100.01–150", "150.01–200", "Over 200"]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
ERRORS = ("#VALUE!", "#N/A", "#DIV/0!", "#REF!", "#NAME?", "#NUM!", "#NULL!", "Err:502", "Err:504", "Err:508")

# ------------------------------------------------------------------ expected values from the brief
EXP_FLAGS = {"flag_invalid": 186, "flag_undefined": 5, "flag_outlier": 241, "flag_zero_rate": 1959,
             "flag_zero_nights": 715, "flag_any_quality": 428, "dup_flag": 32239}
EXP_DQ = {"DQ01": 119390, "DQ02": 119390, "DQ03": 96, "DQ04": 4, "DQ05": 488, "DQ06": 1279,
          "DQ07": "16,340 / 112,593", "DQ08": 5, "DQ09": 1169, "DQ10": 180, "DQ11": 223, "DQ12": 16, "DQ13": 2,
          "DQ14": 1959, "DQ15": 715, "DQ16": 2, "DQ17": 32239, "DQ18": 44224, "DQ19": 119294,
          "DQ20": "7,416 / 14,917"}
EXP_SENS = {"base": 37.04, "any": 37.07, "invund": 37.07, "zn": 37.24, "dup": 27.29, "noshow": 36.03}
EXP_LARGEST = {"hotel_name": 0.07, "lead_band": 0.05, "deposit_type": 0.01, "distribution_channel": 0.03,
               "is_repeated_guest": 0.18, "meal_clean": 0.15, "prev_canc_band": 0.24, "market_segment": 0.90,
               "customer_type": 2.19, "nights_band": 1.33, "adr_band": 1.21, "reserved_room_type": 1.80}
EXP_OV = {  # block -> {label: (bookings, cancelled or None, rate % to 1 dp)}
    "hotel_name": {"City Hotel": (79330, 33102, 41.7), "Resort Hotel": (40060, 11122, 27.8)},
    "lead_band": dict(zip(LEAD, [(19746, 1902, 9.6), (18960, 5283, 27.9), (29553, 11141, 37.7), (26439, 11821, 44.7),
                                 (21544, 11947, 55.5), (3148, 2130, 67.7)])),
    "deposit_type": {"No Deposit": (104641, 29694, 28.4), "Non Refund": (14587, 14494, 99.4),
                     "Refundable": (162, 36, 22.2)},
    "market_segment": {"Online TA": (56477, 20739, 36.7), "Offline TA/TO": (24219, 8311, 34.3),
                       "Groups": (19811, 12097, 61.1), "Direct": (12606, 1934, 15.3), "Corporate": (5295, 992, 18.7),
                       "Complementary": (743, 97, 13.1), "Aviation": (237, 52, 21.9), "Undefined": (2, 2, 100.0)},
    "distribution_channel": {"TA/TO": (97870, 40152, 41.0), "Direct": (14645, 2557, 17.5),
                             "Corporate": (6677, 1474, 22.1), "GDS": (193, 37, 19.2), "Undefined": (5, 4, 80.0)},
    "customer_type": {"Transient": (89613, 36514, 40.7), "Transient-Party": (25124, 6389, 25.4),
                      "Contract": (4076, 1262, 31.0), "Group": (577, 59, 10.2)},
    "is_repeated_guest": {"No": (115580, 43672, 37.8), "Yes": (3810, 552, 14.5)},
    "prev_canc_band": {"0": (112906, 38282, 33.9), "1": (6051, 5714, 94.4), "2+": (433, 228, 52.7)},
    "prev_notcanc_band": {"0": (115770, 44024, 38.0), "1+": (3620, 200, 5.5)},
    "meal_clean": {"BB": (92310, 34510, 37.4), "HB": (14463, 4984, 34.5), "SC": (11819, 4252, 36.0),
                   "FB": (798, 478, 59.9)},
    "changes_band": {"0": (101314, 41391, 40.9), "1": (12701, 1807, 14.2), "2": (3805, 766, 20.1),
                     "3+": (1570, 260, 16.6)},
    "requests_band": {"0": (70318, 33556, 47.7), "1": (33226, 7318, 22.0), "2": (12969, 2866, 22.1),
                      "3+": (2877, 484, 16.8)},
    "waiting_band": {"0": (115692, 41865, 36.2), "1+": (3698, 2359, 63.8)},
    "nights_band": dict(zip(NIGHTS, [(715, 35, 4.9), (21020, 5271, 25.1), (54719, 23514, 43.0), (37679, 13525, 35.9),
                                     (5257, 1879, 35.7)])),
    "adr_band": dict(zip(ADR, [(1960, 213, 10.9), (10037, 2331, 23.2), (55641, 21669, 38.9), (35244, 13724, 38.9),
                               (11579, 4395, 38.0), (4929, 1892, 38.4)])),
    "with_children": {"No": (110058, 40965, 37.2), "Yes": (9332, 3259, 34.9)},
    "arrival_month": dict(zip(MONTHS, [(5929, None, 30.5), (7972, None, 33.3), (9794, None, 32.2), (11089, None, 40.8),
                                       (11791, None, 39.7), (10939, None, 41.5), (12661, None, 37.5),
                                       (13877, None, 37.8), (10508, None, 39.2), (11160, None, 38.0),
                                       (6794, None, 31.2), (6780, None, 35.0)]), **{"Invalid date": (96, None, None)}),
    "arrival_year": {"2022": (21996, None, 37.0), "2023": (56611, None, 35.9), "2024": (40687, None, 38.7),
                     "Invalid date": (96, None, None)},
    "parking_band": {"0": (111974, 44224, 39.5), "1+": (7416, 0, 0.0)},
    "room_changed": {"Same": (104473, 43422, 41.6), "Different": (14917, 802, 5.4)},
    "reserved_room_type": {"A": (85994, None, 39.1), "P": (12, None, 100.0)},
}
BLOCK_ORDER = list(EXP_OV)


def r1(x):
    return Decimal(repr(float(x) * 100)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)


def r2pts(x):
    return Decimal(repr(float(x))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


# ================================================================== stage: recalc (the brief's LibreOffice command)
def stage_recalc(built, outdir):
    os.makedirs(outdir, exist_ok=True)
    t = time.time()
    res = subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", outdir, built],
                         capture_output=True, text=True, timeout=7200)
    secs = time.time() - t
    print(res.stdout[-2000:], res.stderr[-2000:])
    with open(os.path.join(outdir, "recalc_time.txt"), "w") as fh:
        fh.write(f"{secs:.0f}\n")
    print(f"LibreOffice convert (load, recalculate, save): {secs:.0f}s")


# ================================================================== shared category logic (independent of the build)
def band_lead(x):
    if x <= 7:
        return "0–7"
    elif x <= 30:
        return "8–30"
    elif x <= 90:
        return "31–90"
    elif x <= 180:
        return "91–180"
    elif x <= 365:
        return "181–365"
    return "Over 365"


def band_nights(x):
    if x == 0:
        return "0"
    elif x == 1:
        return "1"
    elif x <= 3:
        return "2–3"
    elif x <= 7:
        return "4–7"
    return "8+"


def band_adr(x):
    if x <= 0:
        return "0 or less"
    elif x <= 50:
        return "0.01–50"
    elif x <= 100:
        return "50.01–100"
    elif x <= 150:
        return "100.01–150"
    elif x <= 200:
        return "150.01–200"
    return "Over 200"


def b0123(x):
    return "0" if x == 0 else "1" if x == 1 else "2" if x == 2 else "3+"


# ================================================================== Method B: plain row loop (no pandas)
def method_b(src):
    from openpyxl import load_workbook
    wb = load_workbook(src, read_only=True)
    it = wb["hotels"].iter_rows(values_only=True)
    hdr = list(next(it))
    ix = {h: i for i, h in enumerate(hdr)}
    rows = [list(r) for r in it]
    wb.close()
    adr_max = max(r[ix["Average_Daily_Rate"]] for r in rows)
    seen = set()
    recs = []
    for r in rows:
        g = lambda f: r[ix[f]]
        children = None if g("children") == "NA" else g("children")
        ch0 = 0 if children is None else children
        adults, babies, adr = g("adults"), g("babies"), g("Average_Daily_Rate")
        guests = adults + ch0 + babies
        nights = g("stays_in_weekend_nights") + g("stays_in_week_nights")
        canc = 1 if g("is_canceled") == "Yes" else 0
        valid = isinstance(g("arrival_date"), datetime)
        key = tuple(r[1:])
        dup = 1 if key in seen else 0
        seen.add(key)
        inv = children is None or guests == 0 or adr < 0 or adr == adr_max
        und = g("market_segment") == "Undefined" or g("distribution_channel") == "Undefined"
        out = (adults == 0 and guests > 0) or adults >= 10 or ch0 >= 10 or babies >= 9 or \
            g("required_car_parking_spaces") >= 8
        cats = {
            "hotel_name": g("hotel_name"), "lead_band": band_lead(g("lead_time")), "deposit_type": g("deposit_type"),
            "market_segment": g("market_segment"), "distribution_channel": g("distribution_channel"),
            "customer_type": g("customer_type"), "is_repeated_guest": g("is_repeated_guest"),
            "prev_canc_band": "0" if g("previous_cancellations") == 0 else "1" if g("previous_cancellations") == 1
            else "2+",
            "prev_notcanc_band": "0" if g("previous_bookings_not_canceled") == 0 else "1+",
            "meal_clean": "SC" if g("meal") == "Undefined" else g("meal"),
            "changes_band": b0123(g("booking_changes")), "requests_band": b0123(g("total_of_special_requests")),
            "waiting_band": "0" if g("days_in_waiting_list") == 0 else "1+", "nights_band": band_nights(nights),
            "adr_band": band_adr(adr), "with_children": "Yes" if ch0 + babies > 0 else "No",
            "arrival_month": MONTHS[g("arrival_date").month - 1] if valid else "Invalid date",
            "arrival_year": str(g("arrival_date").year) if valid else "Invalid date",
            "parking_band": "0" if g("required_car_parking_spaces") == 0 else "1+",
            "room_changed": "Same" if g("reserved_room_type") == g("assigned_room_type") else "Different",
            "reserved_room_type": g("reserved_room_type"),
        }
        recs.append(dict(canc=canc, inv=inv, und=und, out=out, any=inv or und or out, zr=adr == 0, zn=nights == 0,
                         dup=dup, valid=valid, cats=cats, raw=r, guests=guests, adults=adults, ch0=ch0, babies=babies,
                         adr=adr, children=children, status=g("reservation_status")))
    return hdr, rows, recs, adr_max


def measures_from_records(recs, hdr, rows, adr_max):
    """Every verified quantity, computed from per-row records with plain loops."""
    ix = {h: i for i, h in enumerate(hdr)}
    M = {}
    n_all = len(recs)
    c_all = sum(x["canc"] for x in recs)
    for f in ("inv", "und", "out", "zr", "zn", "any", "dup"):
        M[("flag", f)] = sum(1 for x in recs if x[f])
    # one-variable tables and the sensitivity detail
    for blk in BLOCK_ORDER:
        agg = defaultdict(lambda: Counter())
        for x in recs:
            a = agg[x["cats"][blk]]
            a["n"] += 1
            a["c"] += x["canc"]
            if x["any"]:
                a["f"] += 1
                a["fc"] += x["canc"]
            a["inv"] += x["inv"]
            a["out"] += x["out"]
            a["und"] += x["und"]
        for lab, a in agg.items():
            M[("ov", blk, lab, "n")] = a["n"]
            M[("ov", blk, lab, "c")] = a["c"]
            M[("ov", blk, lab, "rate")] = a["c"] / a["n"]
            M[("ov", blk, lab, "sh_b")] = a["n"] / n_all
            M[("ov", blk, lab, "sh_c")] = a["c"] / c_all
            for k in ("f", "fc", "inv", "out", "und"):
                M[("s3", blk, lab, k)] = a[k]
            n2 = a["n"] - a["f"]
            M[("s3", blk, lab, "n2")] = n2
            if n2:
                M[("s3", blk, lab, "rate2")] = (a["c"] - a["fc"]) / n2
                M[("s3", blk, lab, "chg")] = ((a["c"] - a["fc"]) / n2 - a["c"] / a["n"]) * 100
        M[("ov_total", blk, "n")] = sum(a["n"] for a in agg.values())
        M[("ov_total", blk, "c")] = sum(a["c"] for a in agg.values())
        counted = [(abs(M[("s3", blk, lab, "chg")]), lab) for lab, a in agg.items()
                   if a["n"] >= 30 and ("s3", blk, lab, "chg") in M]
        best = max(counted)
        M[("s2", blk, "largest")] = best[0]
        M[("s2", blk, "where")] = best[1]
    # sensitivity, part 1

    def rate(sel):
        xs = [x for x in recs if sel(x)]
        return len(xs), sum(x["canc"] for x in xs), sum(x["canc"] for x in xs) / len(xs)
    base = c_all / n_all
    for key, sel in (("base", lambda x: True), ("any", lambda x: not x["any"]),
                     ("invund", lambda x: not (x["inv"] or x["und"])), ("zn", lambda x: not x["zn"]),
                     ("dup", lambda x: not x["dup"])):
        n_, c_, r_ = rate(sel)
        M[("s1", key, "n")], M[("s1", key, "c")], M[("s1", key, "rate")] = n_, c_, r_
        M[("s1", key, "chg")] = (r_ - base) * 100
    canceled = sum(1 for x in recs if x["status"] == "Canceled")
    M[("s1", "noshow", "n")], M[("s1", "noshow", "c")] = n_all, canceled
    M[("s1", "noshow", "rate")] = canceled / n_all
    M[("s1", "noshow", "chg")] = (canceled / n_all - base) * 100
    # data quality log counts
    col = lambda f: [r[ix[f]] for r in rows]
    isc, rep = col("is_canceled"), col("is_repeated_guest")
    M[("dq", "DQ01")] = sum(1 for a, b in zip(isc, rep) if a in ("Yes", "No") and b in ("Yes", "No"))
    M[("dq", "DQ02")] = sum(1 for v in col("arrival_date") if v is not None)
    M[("dq", "DQ03")] = sum(1 for v in col("arrival_date") if v == "29 February 2023")
    M[("dq", "DQ04")] = sum(1 for v in col("children") if v == "NA")
    M[("dq", "DQ05")] = sum(1 for v in col("country") if v == "NULL")
    M[("dq", "DQ06")] = sum(1 for v in col("country") if v == "CN")
    M[("dq", "DQ07")] = f"{sum(1 for v in col('agent') if v == 'NULL'):,} / {sum(1 for v in col('company') if v == 'NULL'):,}"
    M[("dq", "DQ08")] = M[("flag", "und")]
    M[("dq", "DQ09")] = sum(1 for v in col("meal") if v == "Undefined")
    M[("dq", "DQ10")] = sum(1 for x in recs if x["guests"] == 0)
    M[("dq", "DQ11")] = sum(1 for x in recs if x["adults"] == 0 and x["guests"] > 0)
    M[("dq", "DQ12")] = sum(1 for x in recs if x["adults"] >= 10 or x["ch0"] >= 10 or x["babies"] >= 9)
    M[("dq", "DQ13")] = sum(1 for x in recs if x["adr"] < 0 or x["adr"] == adr_max)
    M[("dq", "DQ14")] = M[("flag", "zr")]
    M[("dq", "DQ15")] = M[("flag", "zn")]
    M[("dq", "DQ16")] = sum(1 for v in col("required_car_parking_spaces") if v >= 8)
    M[("dq", "DQ17")] = M[("flag", "dup")]
    M[("dq", "DQ18")] = c_all
    M[("dq", "DQ19")] = sum(1 for x in recs if x["valid"])
    park = [x for x in recs if x["cats"]["parking_band"] == "1+"]
    rc = [x for x in recs if x["cats"]["room_changed"] == "Different"]
    M[("dq", "DQ20")] = f"{len(park):,} / {len(rc):,}"
    # supplementary values (Sensitivity Part 4)
    seg = col("market_segment")
    meal = col("meal")
    sc_raw = [x for x, m in zip(recs, meal) if m == "SC"]
    sc_cl = [x for x in recs if x["cats"]["meal_clean"] == "SC"]
    und_kept = [x for x in recs if not x["und"]]
    adrs = sorted(x["adr"] for x in recs)
    med = (adrs[N // 2 - 1] + adrs[N // 2]) / 2
    pc = col("previous_cancellations")
    dup_go = sum(1 for x, s in zip(recs, seg) if x["dup"] and s in ("Groups", "Offline TA/TO"))
    pc11 = [s for v, s in zip(pc, seg) if 11 <= v <= 26]
    over = [r for r, x in zip(rows, recs) if x["cats"]["lead_band"] == "Over 365"]
    dates = [v for v in col("arrival_date") if isinstance(v, datetime)]
    S = {"und_rate": sum(x["canc"] for x in und_kept) / len(und_kept),
         "sc_raw_n": len(sc_raw), "sc_raw_rate": sum(x["canc"] for x in sc_raw) / len(sc_raw),
         "sc_clean_n": len(sc_cl), "sc_clean_rate": sum(x["canc"] for x in sc_cl) / len(sc_cl),
         "zr_comp": sum(1 for x, s in zip(recs, seg) if x["zr"] and s == "Complementary"),
         "zn_zr": sum(1 for x in recs if x["zn"] and x["zr"]), "adr_median": med, "adr_ratio": adr_max / med,
         "dup_go": dup_go, "dup_go_sh": dup_go / M[("flag", "dup")],
         "go_sh": sum(1 for s in seg if s in ("Groups", "Offline TA/TO")) / n_all,
         "over365_sh": len(over) / n_all,
         "over365_nr": sum(1 for r in over if r[ix["deposit_type"]] == "Non Refund") / len(over),
         "pc11": len(pc11), "pc11_go": sum(1 for s in pc11 if s in ("Groups", "Offline TA/TO")),
         "pc11_go_sh": sum(1 for s in pc11 if s in ("Groups", "Offline TA/TO")) / len(pc11),
         "park_n": len(park), "park_c": sum(x["canc"] for x in park), "rc_n": len(rc),
         "rc_rate": sum(x["canc"] for x in rc) / len(rc), "date_min": min(dates), "date_max": max(dates),
         "feb_valid": sum(1 for d in dates if d.month == 2)}
    S["und_chg"] = (S["und_rate"] - base) * 100
    for k, v in S.items():
        M[("supp", k)] = v
    return M


# ================================================================== Method A: pandas
def method_a(src):
    import numpy as np
    import pandas as pd
    raw = pd.read_excel(src, sheet_name="hotels", engine="openpyxl", keep_default_na=False, dtype=object)
    d = pd.DataFrame(index=raw.index)
    num = lambda c: pd.to_numeric(raw[c].replace({"NA": np.nan, "NULL": np.nan}))
    d["canc"] = (raw["is_canceled"] == "Yes").astype(int)
    children = num("children")
    ch0 = children.fillna(0)
    adults, babies, adr = num("adults"), num("babies"), num("Average_Daily_Rate")
    guests = adults + ch0 + babies
    nights = num("stays_in_weekend_nights") + num("stays_in_week_nights")
    parking = num("required_car_parking_spaces")
    valid = raw["arrival_date"].map(lambda v: isinstance(v, datetime))
    d["inv"] = children.isna() | (guests == 0) | (adr < 0) | (adr == adr.max())
    d["und"] = (raw["market_segment"] == "Undefined") | (raw["distribution_channel"] == "Undefined")
    d["out"] = ((adults == 0) & (guests > 0)) | (adults >= 10) | (ch0 >= 10) | (babies >= 9) | (parking >= 8)
    d["any"] = d["inv"] | d["und"] | d["out"]
    d["zr"], d["zn"] = adr == 0, nights == 0
    d["dup"] = raw.duplicated(subset=list(raw.columns[1:]), keep="first")
    lt = num("lead_time")
    cuts = lambda s, bins, labels: pd.cut(s, bins=bins, labels=labels, right=True).astype(str)
    cats = {
        "hotel_name": raw["hotel_name"], "lead_band": cuts(lt, [-1, 7, 30, 90, 180, 365, 1e9], LEAD),
        "deposit_type": raw["deposit_type"], "market_segment": raw["market_segment"],
        "distribution_channel": raw["distribution_channel"], "customer_type": raw["customer_type"],
        "is_repeated_guest": raw["is_repeated_guest"],
        "prev_canc_band": cuts(num("previous_cancellations"), [-1, 0, 1, 1e9], ["0", "1", "2+"]),
        "prev_notcanc_band": cuts(num("previous_bookings_not_canceled"), [-1, 0, 1e9], ["0", "1+"]),
        "meal_clean": raw["meal"].replace({"Undefined": "SC"}),
        "changes_band": cuts(num("booking_changes"), [-1, 0, 1, 2, 1e9], ["0", "1", "2", "3+"]),
        "requests_band": cuts(num("total_of_special_requests"), [-1, 0, 1, 2, 1e9], ["0", "1", "2", "3+"]),
        "waiting_band": cuts(num("days_in_waiting_list"), [-1, 0, 1e9], ["0", "1+"]),
        "nights_band": cuts(nights, [-1, 0, 1, 3, 7, 1e9], NIGHTS),
        "adr_band": cuts(adr, [-1e9, 0, 50, 100, 150, 200, 1e9], ADR),
        "with_children": np.where(ch0 + babies > 0, "Yes", "No"),
        "arrival_month": raw["arrival_date"].map(lambda v: MONTHS[v.month - 1] if isinstance(v, datetime)
                                                 else "Invalid date"),
        "arrival_year": raw["arrival_date"].map(lambda v: str(v.year) if isinstance(v, datetime) else "Invalid date"),
        "parking_band": np.where(parking == 0, "0", "1+"),
        "room_changed": np.where(raw["reserved_room_type"] == raw["assigned_room_type"], "Same", "Different"),
        "reserved_room_type": raw["reserved_room_type"],
    }
    M = {}
    n_all, c_all = len(d), int(d["canc"].sum())
    for f in ("inv", "und", "out", "zr", "zn", "any", "dup"):
        M[("flag", f)] = int(d[f].sum())
    for blk in BLOCK_ORDER:
        g = pd.DataFrame({"k": cats[blk], "c": d["canc"], "f": d["any"], "inv": d["inv"], "out": d["out"],
                          "und": d["und"]})
        g["fc"] = g["f"] & (g["c"] == 1)
        t = g.groupby("k").agg(n=("c", "size"), c=("c", "sum"), f=("f", "sum"), fc=("fc", "sum"), inv=("inv", "sum"),
                               out=("out", "sum"), und=("und", "sum"))
        best = []
        for lab, row in t.iterrows():
            n_, c_ = int(row["n"]), int(row["c"])
            M[("ov", blk, lab, "n")], M[("ov", blk, lab, "c")] = n_, c_
            M[("ov", blk, lab, "rate")] = c_ / n_
            M[("ov", blk, lab, "sh_b")], M[("ov", blk, lab, "sh_c")] = n_ / n_all, c_ / c_all
            for k in ("f", "fc", "inv", "out", "und"):
                M[("s3", blk, lab, k)] = int(row[k])
            n2 = n_ - int(row["f"])
            M[("s3", blk, lab, "n2")] = n2
            if n2:
                r2 = (c_ - int(row["fc"])) / n2
                M[("s3", blk, lab, "rate2")] = r2
                M[("s3", blk, lab, "chg")] = (r2 - c_ / n_) * 100
                if n_ >= 30:
                    best.append((abs(M[("s3", blk, lab, "chg")]), lab))
        M[("ov_total", blk, "n")], M[("ov_total", blk, "c")] = int(t["n"].sum()), int(t["c"].sum())
        b = max(best)
        M[("s2", blk, "largest")], M[("s2", blk, "where")] = b[0], b[1]
    base = c_all / n_all
    for key, mask in (("base", pd.Series(True, index=d.index)), ("any", ~d["any"]), ("invund", ~(d["inv"] | d["und"])),
                      ("zn", ~d["zn"]), ("dup", ~d["dup"])):
        M[("s1", key, "n")], M[("s1", key, "c")] = int(mask.sum()), int(d["canc"][mask].sum())
        M[("s1", key, "rate")] = d["canc"][mask].sum() / mask.sum()
        M[("s1", key, "chg")] = (M[("s1", key, "rate")] - base) * 100
    canceled = int((raw["reservation_status"] == "Canceled").sum())
    M[("s1", "noshow", "n")], M[("s1", "noshow", "c")] = n_all, canceled
    M[("s1", "noshow", "rate")] = canceled / n_all
    M[("s1", "noshow", "chg")] = (canceled / n_all - base) * 100
    yn = lambda c: raw[c].isin(["Yes", "No"])
    M[("dq", "DQ01")] = int((yn("is_canceled") & yn("is_repeated_guest")).sum())
    M[("dq", "DQ02")] = int((raw["arrival_date"] != "").sum())
    M[("dq", "DQ03")] = int((raw["arrival_date"] == "29 February 2023").sum())
    M[("dq", "DQ04")] = int((raw["children"] == "NA").sum())
    M[("dq", "DQ05")] = int((raw["country"] == "NULL").sum())
    M[("dq", "DQ06")] = int((raw["country"] == "CN").sum())
    M[("dq", "DQ07")] = f"{int((raw['agent'] == 'NULL').sum()):,} / {int((raw['company'] == 'NULL').sum()):,}"
    M[("dq", "DQ08")] = M[("flag", "und")]
    M[("dq", "DQ09")] = int((raw["meal"] == "Undefined").sum())
    M[("dq", "DQ10")] = int((guests == 0).sum())
    M[("dq", "DQ11")] = int(((adults == 0) & (guests > 0)).sum())
    M[("dq", "DQ12")] = int(((adults >= 10) | (ch0 >= 10) | (babies >= 9)).sum())
    M[("dq", "DQ13")] = int(((adr < 0) | (adr == adr.max())).sum())
    M[("dq", "DQ14")] = M[("flag", "zr")]
    M[("dq", "DQ15")] = M[("flag", "zn")]
    M[("dq", "DQ16")] = int((parking >= 8).sum())
    M[("dq", "DQ17")] = M[("flag", "dup")]
    M[("dq", "DQ18")] = c_all
    M[("dq", "DQ19")] = int(valid.sum())
    park, rc = parking >= 1, raw["reserved_room_type"] != raw["assigned_room_type"]
    M[("dq", "DQ20")] = f"{int(park.sum()):,} / {int(rc.sum()):,}"
    seg = raw["market_segment"]
    go = seg.isin(["Groups", "Offline TA/TO"])
    pcn = num("previous_cancellations")
    pc11 = (pcn >= 11) & (pcn <= 26)
    over = cats["lead_band"] == "Over 365"
    dts = raw["arrival_date"][valid]
    S = {"und_rate": d["canc"][~d["und"]].mean(), "sc_raw_n": int((raw["meal"] == "SC").sum()),
         "sc_raw_rate": d["canc"][raw["meal"] == "SC"].mean(),
         "sc_clean_n": int((cats["meal_clean"] == "SC").sum()),
         "sc_clean_rate": d["canc"][cats["meal_clean"] == "SC"].mean(),
         "zr_comp": int((d["zr"] & (seg == "Complementary")).sum()), "zn_zr": int((d["zn"] & d["zr"]).sum()),
         "adr_median": float(adr.median()), "adr_ratio": float(adr.max() / adr.median()),
         "dup_go": int((d["dup"] & go).sum()), "dup_go_sh": (d["dup"] & go).sum() / d["dup"].sum(),
         "go_sh": go.mean(), "over365_sh": over.mean(),
         "over365_nr": (raw["deposit_type"][over] == "Non Refund").mean(), "pc11": int(pc11.sum()),
         "pc11_go": int((pc11 & go).sum()), "pc11_go_sh": (pc11 & go).sum() / pc11.sum(),
         "park_n": int(park.sum()), "park_c": int(d["canc"][park].sum()), "rc_n": int(rc.sum()),
         "rc_rate": d["canc"][rc].mean(), "date_min": min(dts), "date_max": max(dts),
         "feb_valid": int(sum(1 for v in dts if v.month == 2))}
    S["und_chg"] = (S["und_rate"] - base) * 100
    for k, v in S.items():
        M[("supp", k)] = v
    return M, raw


# ================================================================== Method C: the workbook's formulas
def find_row(ws, prefix, col=1, start=1, stop=3000):
    for r in range(start, stop):
        v = ws.cell(row=r, column=col).value
        if isinstance(v, str) and v.startswith(prefix):
            return r
    raise KeyError(prefix)


SUPP_LABELS = {  # Sensitivity Part 4 label prefix -> key
    "Overall rate excluding flag_undefined": "und_rate", "Change vs baseline, excluding flag_undefined": "und_chg",
    "meal = SC in the original column (bookings)": "sc_raw_n", "meal = SC in the original column (rate)": "sc_raw_rate",
    "meal_clean = SC (bookings)": "sc_clean_n", "meal_clean = SC (rate)": "sc_clean_rate",
    "Zero-rate bookings that are Complementary": "zr_comp", "Zero-night stays that also have a rate of 0": "zn_zr",
    "Median Average_Daily_Rate": "adr_median", "Maximum rate ÷ median rate": "adr_ratio",
    "Duplicates that are Groups or Offline TA/TO": "dup_go", "…as a share of all duplicates": "dup_go_sh",
    "Groups + Offline TA/TO as a share of all bookings": "go_sh", "Over 365 band: share of all bookings": "over365_sh",
    "Over 365 band: share that is Non Refund": "over365_nr", "Bookings with 11 to 26 previous cancellations": "pc11",
    "…of which Groups or Offline TA/TO": "pc11_go", "…as a share of those bookings": "pc11_go_sh",
    "Bookings with 1+ parking spaces": "park_n", "…of which cancelled": "park_c",
    "Bookings with a changed room type": "rc_n", "…their cancellation rate": "rc_rate",
    "Earliest valid arrival date": "date_min", "Latest valid arrival date": "date_max",
    "February arrivals with a valid date": "feb_valid",
}


def method_c(recalc):
    from openpyxl import load_workbook
    wb = load_workbook(recalc, data_only=True, read_only=False)
    M, extra = {}, {}
    ov, se, dq, ck = wb["OneVar_Summary"], wb["Sensitivity"], wb["Data_Quality_Log"], wb["Checks"]
    r = 5
    blocks_seen = []
    while r < ov.max_row:
        v = ov.cell(row=r, column=1).value
        if isinstance(v, str) and v.startswith("Block "):
            blk = v.split(": ", 1)[1].split(" ")[0].rstrip(":")
            blocks_seen.append(blk)
            rr = r + 2
            while ov.cell(row=rr, column=1).value != "Total":
                lab = str(ov.cell(row=rr, column=1).value)
                vals = [ov.cell(row=rr, column=c).value for c in range(2, 8)]
                M[("ov", blk, lab, "n")], M[("ov", blk, lab, "c")], M[("ov", blk, lab, "rate")] = vals[0], vals[1], vals[2]
                M[("ov", blk, lab, "sh_b")], M[("ov", blk, lab, "sh_c")] = vals[3], vals[4]
                extra[("warn", blk, lab)] = vals[5]
                rr += 1
            M[("ov_total", blk, "n")] = ov.cell(row=rr, column=2).value
            M[("ov_total", blk, "c")] = ov.cell(row=rr, column=3).value
            extra[("stands", blk)] = ov.cell(row=rr + 1, column=1).value
            r = rr + 1
        else:
            r += 1
    extra["blocks"] = blocks_seen
    s1 = find_row(se, "Part 1") + 2
    for i, key in enumerate(("base", "any", "invund", "zn", "dup", "noshow")):
        rr = s1 + i
        M[("s1", key, "n")], M[("s1", key, "c")] = se.cell(row=rr, column=2).value, se.cell(row=rr, column=3).value
        M[("s1", key, "rate")], M[("s1", key, "chg")] = se.cell(row=rr, column=4).value, se.cell(row=rr, column=5).value
        extra[("s1match", key)] = se.cell(row=rr, column=7).value
    s2 = find_row(se, "Part 2") + 2
    rr = s2
    while se.cell(row=rr, column=1).value not in (None, ""):
        blk = se.cell(row=rr, column=1).value
        M[("s2", blk, "largest")] = se.cell(row=rr, column=3).value
        M[("s2", blk, "where")] = str(se.cell(row=rr, column=4).value)
        extra[("s2", blk)] = dict(exp=se.cell(row=rr, column=6).value, match=se.cell(row=rr, column=7).value,
                                  why=se.cell(row=rr, column=8).value, signed=se.cell(row=rr, column=5).value)
        rr += 1
    s3 = find_row(se, "Part 3") + 2
    rr = s3
    while se.cell(row=rr, column=1).value not in (None, ""):
        blk, lab = se.cell(row=rr, column=1).value, str(se.cell(row=rr, column=2).value)
        for k, c in (("f", 6), ("inv", 7), ("out", 8), ("und", 9), ("fc", 10), ("n2", 11), ("rate2", 12), ("chg", 13)):
            v = se.cell(row=rr, column=c).value
            if v != "n/a":
                M[("s3", blk, lab, k)] = v
        rr += 1
    s4 = find_row(se, "Part 4") + 2
    rr = s4
    while se.cell(row=rr, column=1).value not in (None, ""):
        lab = se.cell(row=rr, column=1).value
        key = [k for p, k in SUPP_LABELS.items() if lab.startswith(p)]
        assert len(key) == 1, lab
        M[("supp", key[0])] = se.cell(row=rr, column=2).value
        rr += 1
    for rr in range(5, 25):
        did = dq.cell(row=rr, column=1).value
        M[("dq", did)] = dq.cell(row=rr, column=5).value
        extra[("dq", did)] = dict(exp=dq.cell(row=rr, column=6).value, match=dq.cell(row=rr, column=7).value,
                                  row=[dq.cell(row=rr, column=c).value for c in range(1, 12)])
    cl = wb["Clean_Data"]
    hdr = [c.value for c in cl[1]]
    extra["clean_hdr"] = hdr
    extra["banner"], extra["banner_count"] = ck["A3"].value, ck["A4"].value
    checks = []
    rr = 7
    while ck.cell(row=rr, column=2).value not in (None, ""):
        checks.append([ck.cell(row=rr, column=c).value for c in range(1, 6)])
        rr += 1
    extra["checks"] = checks
    extra["checks_info"] = ck.cell(row=rr + 1, column=1).value
    pf = wb["Column_Profile"]
    prof = {}
    for rr in range(5, 34):
        prof[pf.cell(row=rr, column=2).value] = [pf.cell(row=rr, column=c).value for c in range(1, 15)]
    extra["profile"] = prof
    # every displayed number on the summary sheets, for the Word number check
    shown = set()
    for name in ("README", "Column_Profile", "Data_Quality_Log", "Sensitivity", "OneVar_Summary", "Checks"):
        ws = wb[name]
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None:
                    continue
                for s in displayed(v, c.number_format):
                    shown.add(s)
    extra["shown_text"] = shown
    wb.close()
    # flag sums straight from the recalculated tblClean
    wb2 = load_workbook(recalc, data_only=True, read_only=True)
    ws = wb2["Clean_Data"]
    idx = {h: i for i, h in enumerate(hdr)}
    sums = Counter()
    err = 0
    none_in_formula_cols = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        for f in ("flag_invalid", "flag_undefined", "flag_outlier", "flag_zero_rate", "flag_zero_nights",
                  "flag_any_quality", "dup_flag"):
            sums[f] += row[idx[f]]
        for v in row[29:68]:
            if v is None:
                none_in_formula_cols += 1
            elif isinstance(v, str) and v in ERRORS:
                err += 1
    wb2.close()
    for f, k in (("flag_invalid", "inv"), ("flag_undefined", "und"), ("flag_outlier", "out"),
                 ("flag_zero_rate", "zr"), ("flag_zero_nights", "zn"), ("flag_any_quality", "any"),
                 ("dup_flag", "dup")):
        M[("flag", k)] = sums[f]
    extra["clean_errors"], extra["clean_none"] = err, none_in_formula_cols
    return M, extra


def displayed(v, fmt):
    """Strings a reader could see for a cell value (its displayed text and the numbers inside it)."""
    out = []
    if isinstance(v, bool):
        return out
    if isinstance(v, (int, float)):
        out += [f"{v:,.0f}", f"{v:,.1f}", f"{v:,.2f}", f"{v * 100:.1f}%", f"{v * 100:.2f}%", f"{v * 100:.0f}%",
                f"{v:+.2f}", f"{v:.2f}", f"{v:.1f}", f"{v * 100:.1f}", f"{v * 100:.2f}"]
    elif isinstance(v, datetime):
        out += [f"{v.day} {v:%B %Y}"]
    else:
        out += re.findall(r"[−-]?\d[\d,]*(?:\.\d+)?%?", str(v))
    return out


# ================================================================== structural checks on the built file
def structure(src, built, recalc, rec):
    from openpyxl import load_workbook
    sw = load_workbook(src, read_only=True)
    src_rows = list(sw["hotels"].iter_rows(values_only=True))
    sw.close()
    bw = load_workbook(built, read_only=True)
    names = bw.sheetnames
    want = ["README", "Raw_Data", "Column_Profile", "Clean_Data", "Data_Quality_Log", "Sensitivity", "OneVar_Summary",
            "Checks"]
    rec("Structure", "Sheet order", "", "", ", ".join(names), ", ".join(want), names == want)
    raw_rows = list(bw["Raw_Data"].iter_rows(values_only=True))
    exact = len(raw_rows) == len(src_rows)
    diffs = 0
    for a, b in zip(raw_rows, src_rows):
        for x, y in zip(a[:29], b):
            if not (x == y and type(x) is type(y)):
                diffs += 1
    nulls = Counter()
    for r in raw_rows[1:]:
        for j, tok in ((8, "NA"), (11, "NULL"), (21, "NULL"), (22, "NULL")):
            nulls[(j, tok)] += r[j] == tok
    texts = sum(1 for r in raw_rows[1:] if r[4] == "29 February 2023")
    rec("Raw_Data integrity", "Raw_Data equals the source cell for cell (value AND type), 119,391 × 29 incl. header",
        "", "", f"{diffs} differing cells; {len(raw_rows)} rows", "0 differing cells",
        exact and diffs == 0)
    rec("Raw_Data integrity", 'Raw_Data keeps the texts "NULL" / "NA" / "29 February 2023"', "", "",
        f"country NULL {nulls[(11, 'NULL')]:,}, agent NULL {nulls[(21, 'NULL')]:,}, company NULL "
        f"{nulls[(22, 'NULL')]:,}, children NA {nulls[(8, 'NA')]:,}, text dates {texts}",
        "488 / 16,340 / 112,593 / 4 / 96",
        (nulls[(11, "NULL")], nulls[(21, "NULL")], nulls[(22, "NULL")], nulls[(8, "NA")], texts) ==
        (488, 16340, 112593, 4, 96))
    clean = bw["Clean_Data"]
    cdiff, cbad, formula_bad, dup_sum = 0, 0, 0, 0
    hdr = None
    for i, row in enumerate(clean.iter_rows(values_only=True)):
        if i == 0:
            hdr = list(row)
            continue
        rr = raw_rows[i]
        for j in range(29):
            if not (row[j] == rr[j] and type(row[j]) is type(rr[j])):
                cdiff += 1
                if not (row[j] is None and rr[j] in ("NULL", "NA") and j in (8, 11, 21, 22)):
                    cbad += 1
        fs = row[29:67]
        if not all(isinstance(x, str) and x.startswith("=") and re.search(rf"\D{i + 1}(\D|$)", x) for x in fs):
            formula_bad += 1
        dup_sum += row[67]
    rec("Raw_Data integrity", "Clean_Data's first 29 columns differ from Raw_Data in exactly 129,425 cells", "", "",
        f"{cdiff:,} differing cells, {cbad} of them NOT a NULL/NA → blank conversion", "129,425 / 0",
        cdiff == 129425 and cbad == 0)
    rec("Structure", "Clean_Data: 38 derived columns are one formula per row, every row", "", "",
        f"{formula_bad} rows not formulas", 0, formula_bad == 0)
    rec("Structure", "Clean_Data: dup_flag value column sums to 32,239", "", "", dup_sum, 32239, dup_sum == 32239)
    rec("Structure", "Clean_Data columns", "", "", len(hdr), 68, len(hdr) == 68)
    bw.close()
    fw = load_workbook(built)
    t = fw["Clean_Data"].tables.get("tblClean")
    rec("Structure", "Excel Table tblClean covers A1:BP119391", "", "", t.ref if t else None, "A1:BP119391",
        t is not None and t.ref == "A1:BP119391")
    rngs = {k: v.attr_text for k, v in fw.defined_names.items()}
    bad_rng = [k for k, v in rngs.items() if k.startswith("rng_") and not re.fullmatch(
        r"Clean_Data!\$[A-Z]+\$2:\$[A-Z]+\$119391", v)]
    rec("Structure", "Named ranges rng_* cover exactly rows 2:119391", "", "",
        f"{sum(1 for k in rngs if k.startswith('rng_'))} ranges; {len(bad_rng)} wrong", "0 wrong", not bad_rng)
    rec("Structure", "Named cell ADR_MAX", "", "", rngs.get("ADR_MAX"), "README cell", "ADR_MAX" in rngs)
    banned = re.compile(r"\b(IFS|LET|LAMBDA|XLOOKUP|XMATCH|FILTER|UNIQUE|SORT|SORTBY|SEQUENCE|MAXIFS|MINIFS|"
                        r"TEXTJOIN|CONCAT|SWITCH)\s*\(")
    n_f = n_b = n_s = 0
    typed = []
    for ws in fw.worksheets:
        if ws.title in ("Raw_Data",):
            continue
        rows = ws.iter_rows(min_row=1, max_row=3) if ws.title == "Clean_Data" else ws.iter_rows()
        for row in rows:
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    n_f += 1
                    n_b += bool(banned.search(c.value))
                    n_s += "[@" in c.value or "tblClean[" in c.value
    for c in fw["OneVar_Summary"]["B"] + fw["OneVar_Summary"]["C"] + fw["OneVar_Summary"]["D"]:
        if isinstance(c.value, (int, float)) and not isinstance(c.value, bool):
            typed.append(c.coordinate)
    rec("Structure", f"Banned (post-2019 / dynamic-array) functions in {n_f:,} formulas checked", "", "", n_b, 0,
        n_b == 0)
    rec("Structure", "Structured [@col] references", "", "", n_s, 0, n_s == 0)
    rec("Structure", "Typed numbers in OneVar_Summary count/rate columns", "", "", ", ".join(typed) or "none", "none",
        not typed)
    rec("Structure", "fullCalcOnLoad set", "", "", fw.calculation.fullCalcOnLoad, True,
        bool(fw.calculation.fullCalcOnLoad))
    rec("Structure", "Raw_Data protected without a password", "", "", fw["Raw_Data"].protection.sheet, True,
        fw["Raw_Data"].protection.sheet and not fw["Raw_Data"].protection.password)
    rec("Structure", "Raw_Data header row frozen", "", "", fw["Raw_Data"].freeze_panes, "A2",
        fw["Raw_Data"].freeze_panes == "A2")
    country_block = any(isinstance(c.value, str) and c.value.startswith("Block") and "country" in c.value
                        for c in fw["OneVar_Summary"]["A"])
    rec("Structure", "No country block in OneVar_Summary (scenario rule)", "", "", country_block, False,
        not country_block)
    fw.close()


# ================================================================== Word number check
def docx_text(path):
    from docx import Document
    d = Document(path)
    parts = [p.text for p in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            seen = set()
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                parts.append(c.text)
    return "\n".join(parts)


CITATION_NUMBERS = {  # numbers that are citations, references, dates of sources or labels, not data values
    "2019", "2019a", "2019b", "2019c", "57", "58", "57–58", "63", "45", "22", "41", "49", "41–49", "10.1016",
    "2018.11.126", "2018", "11", "126", "6.4", "1.", "2.", "3.", "4.", "5.", "6.", "7.",
}


def check_docx_numbers(docx_path, shown, viewed_tokens):
    text = docx_text(docx_path)
    scan = re.sub(r"<https?://[^>]+>", " ", text)  # URLs and DOIs are not data values
    tokens = [t.rstrip(",") for t in re.findall(r"[−-]?\d[\d,]*(?:\.\d+)?%?", scan)]
    unmatched = []
    kinds = Counter()
    for tok in tokens:
        bare = tok.lstrip("−-")
        if tok in shown or bare in shown or bare.rstrip("%") in shown:
            kinds["workbook value"] += 1
        elif bare in CITATION_NUMBERS or tok in viewed_tokens or bare in viewed_tokens:
            kinds["citation, date or section label"] += 1
        else:
            unmatched.append(tok)
    return tokens, kinds, unmatched, text


# ================================================================== stage: check
def pdf_checks(pdf, text):
    """Words broken mid-word (a PDF word that is not a word of the Word text) and words outside the margins."""
    import html
    import unicodedata
    bbox = subprocess.run(["pdftotext", "-bbox", pdf, "-"], capture_output=True, text=True).stdout
    pages = re.findall(r'<page width="([\d.]+)" height="([\d.]+)">(.*?)</page>', bbox, flags=re.S)
    doc_words = set(re.findall(r"[A-Za-z]+", unicodedata.normalize("NFKC", text)))
    broken, outside = [], []
    for pi, (w, h, body) in enumerate(pages, 1):
        left, right = 2.54 / 2.54 * 72 - 1.5, float(w) - 72 + 1.5
        for xmin, ymin, xmax, ymax, word in re.findall(
                r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', body):
            word = unicodedata.normalize("NFKC", html.unescape(word))
            for part in re.findall(r"[A-Za-z]+", word):
                if part not in doc_words:
                    broken.append(f"p{pi}:{part}")
            if float(xmin) < left or float(xmax) > right:
                outside.append(f"p{pi}:{word}")
    return len(pages), broken, outside


def stage_check(src, built, recalc, docx_path, pdf_path, report):
    results = []
    t0 = time.time()

    def rec(section, name, a, b, c, exp, ok):
        results.append((section, name, a, b, c, exp, "PASS" if ok else "FAIL"))

    print("Method A (pandas) ...", flush=True)
    A, rawdf = method_a(src)
    print("Method B (plain row loop) ...", flush=True)
    hdr, rows, recs, adr_max = method_b(src)
    B = measures_from_records(recs, hdr, rows, adr_max)
    print("Method C (recalculated workbook) ...", flush=True)
    Cv, X = method_c(recalc)

    def fmt(v):
        if isinstance(v, float):
            return f"{v:.6f}".rstrip("0").rstrip(".")
        if isinstance(v, datetime):
            return f"{v:%d/%m/%Y}"
        return str(v)

    def same(a, b, c):
        if isinstance(a, (int,)) and not isinstance(a, bool) and isinstance(b, int):
            return a == b == c
        if isinstance(a, float) or isinstance(b, float):
            try:
                return abs(float(a) - float(b)) < 1e-12 and abs(float(a) - float(c)) < 1e-12
            except (TypeError, ValueError):
                return False
        if hasattr(a, "to_pydatetime"):
            a = a.to_pydatetime()
        return a == b == c

    # ---- 1. three-way agreement on every quantity
    keys = sorted(set(A) | set(B), key=str)
    missing = []
    for k in keys:
        a, b, c = A.get(k), B.get(k), Cv.get(k)
        if c is None and k[0] != "s3":
            missing.append(k)
        section = {"ov": "One-variable tables (OneVar_Summary)", "ov_total": "One-variable tables (OneVar_Summary)",
                   "s1": "Sensitivity Part 1", "s2": "Sensitivity Part 2", "s3": "Sensitivity Part 3",
                   "dq": "Data_Quality_Log rows affected", "supp": "Sensitivity Part 4 (values quoted in the notes)",
                   "flag": "Flag sums"}[k[0]]
        if isinstance(c, datetime) and hasattr(a, "to_pydatetime"):
            a = a.to_pydatetime()
        rec(f"Three-way agreement: {section}", " | ".join(map(str, k[1:])), fmt(a), fmt(b), fmt(c), "", same(a, b, c))
    rec("Three-way agreement: coverage", "Quantities computed by A and B but absent from the workbook", "", "",
        ", ".join(map(str, missing)) or "none", "none", not missing)

    # ---- 2. expected values from the brief
    def exp_rec(section, name, key, want, kind="count"):
        vals = [A.get(key), B.get(key), Cv.get(key)]
        if kind == "count":
            ok = all(v == want for v in vals)
            shown_ = [fmt(v) for v in vals]
        elif kind == "pct1":
            ok = all(v is not None and r1(v) == Decimal(str(want)).quantize(Decimal("0.1")) for v in vals)
            shown_ = [f"{r1(v)}%" if v is not None else "None" for v in vals]
        elif kind == "pct2":
            ok = all(v is not None and abs(float(v) * 100 - want) <= 0.005 + 1e-12 for v in vals)
            shown_ = [f"{float(v) * 100:.4f}%" for v in vals]
        elif kind == "pts2":
            ok = all(v is not None and r2pts(v) == Decimal(str(want)).quantize(Decimal("0.01")) for v in vals)
            shown_ = [f"{float(v):.4f}" for v in vals]
        rec(f"Expected values (brief): {section}", name, *shown_, want, ok)

    for f, k in (("flag_invalid", "inv"), ("flag_undefined", "und"), ("flag_outlier", "out"),
                 ("flag_zero_rate", "zr"), ("flag_zero_nights", "zn"), ("flag_any_quality", "any"),
                 ("dup_flag", "dup")):
        exp_rec("flag sums (A3)", f, ("flag", k), EXP_FLAGS[f])
    for did, want in EXP_DQ.items():
        exp_rec("Data_Quality_Log (A4)", did, ("dq", did), want)
        x = X[("dq", did)]
        rec("Expected values (brief): Data_Quality_Log (A4)", f"{did} typed Expected cell and Match?", "", "",
            f"{x['exp']} → {x['match']}", f"{want} → ✔ OK", x["exp"] == want and x["match"] == OK)
    for key, want in EXP_SENS.items():
        exp_rec("Sensitivity overall rates (A5)", key, ("s1", key, "rate"), want, "pct2")
    exp_rec("Sensitivity overall rates (A5)", "any: change +0.03 points", ("s1", "any", "chg"), 0.03, "pts2")
    for blk, want in EXP_LARGEST.items():
        vals = [A[("s2", blk, "largest")], B[("s2", blk, "largest")], Cv[("s2", blk, "largest")]]
        ok = all(r2pts(v) == Decimal(str(want)).quantize(Decimal("0.01")) for v in vals)
        agree = abs(vals[0] - vals[1]) < 1e-12 and abs(vals[0] - vals[2]) < 1e-12
        if not ok and agree and blk in BRIEF_SLIPS:  # documented in DEVIATIONS; shown, never forced to PASS
            results.append(("Expected values (brief): largest change per table (A5)", blk,
                            *[f"{v:.4f}" for v in vals], want, "DIFFERS FROM BRIEF"))
            continue
        rec("Expected values (brief): largest change per table (A5)", blk, *[f"{v:.4f}" for v in vals], want, ok)
    for blk, cats in EXP_OV.items():
        for lab, (nb, nc, rate) in cats.items():
            exp_rec("one-variable tables (A6)", f"{blk} | {lab} | bookings", ("ov", blk, lab, "n"), nb)
            if nc is not None:
                exp_rec("one-variable tables (A6)", f"{blk} | {lab} | cancelled", ("ov", blk, lab, "c"), nc)
            if rate is not None:
                exp_rec("one-variable tables (A6)", f"{blk} | {lab} | rate", ("ov", blk, lab, "rate"), rate, "pct1")
        exp_rec("one-variable tables (A6)", f"{blk} | Total bookings", ("ov_total", blk, "n"), 119390)
        exp_rec("one-variable tables (A6)", f"{blk} | Total cancelled", ("ov_total", blk, "c"), 44224)
    rec("Expected values (brief): one-variable tables (A6)", "Blocks present, in order, no country block", "", "",
        ", ".join(X["blocks"]), ", ".join(BLOCK_ORDER), X["blocks"] == BLOCK_ORDER)
    for blk, lab in (("market_segment", "Undefined"), ("distribution_channel", "Undefined"),
                     ("reserved_room_type", "L"), ("reserved_room_type", "P")):
        rec("Expected values (brief): one-variable tables (A6)", f"n < 30 warning on {blk} | {lab}", "", "",
            X[("warn", blk, lab)], "n < 30: unreliable", X[("warn", blk, lab)] == "n < 30: unreliable")
    n_warn = sum(1 for k, v in X.items() if isinstance(k, tuple) and k[0] == "warn" and v)
    rec("Expected values (brief): one-variable tables (A6)", "Only categories with n < 30 carry the warning", "", "",
        n_warn, 4, n_warn == 4)

    # ---- profile facts (A1) on the Column_Profile sheet
    P = X["profile"]
    prof_checks = [
        ("lead_time", "min / median / mean / max / zeros", (P["lead_time"][6], P["lead_time"][7],
                                                             round(P["lead_time"][8], 2), P["lead_time"][9],
                                                             P["lead_time"][10]), (0, 69, 104.01, 737, 6345)),
        ("adults", "zeros / max", (P["adults"][10], P["adults"][9]), (403, 55)),
        ("Average_Daily_Rate", "min / median / max / zeros", (P["Average_Daily_Rate"][6], P["Average_Daily_Rate"][7],
                                                              P["Average_Daily_Rate"][9], P["Average_Daily_Rate"][10]),
         (-6.38, 94.575, 5400, 1959)),
        ("children", "text values / zeros", (P["children"][12], P["children"][10]), (4, 110796)),
        ("arrival_date", "first / last date", (P["arrival_date"][6], P["arrival_date"][9]),
         (datetime(2022, 7, 1), datetime(2024, 8, 31))),
        ("country", "distinct values", (P["country"][5],), (178,)),
        ("booking_order_id", "distinct values", (P["booking_order_id"][5],), (119390,)),
        ("total_of_special_requests", "min / max", (P["total_of_special_requests"][6],
                                                    P["total_of_special_requests"][9]), (0, 5)),
        ("booking_changes", "min / max", (P["booking_changes"][6], P["booking_changes"][9]), (0, 21)),
        ("days_in_waiting_list", "min / max", (P["days_in_waiting_list"][6], P["days_in_waiting_list"][9]), (0, 391)),
        ("previous_cancellations", "min / max", (P["previous_cancellations"][6], P["previous_cancellations"][9]),
         (0, 26)),
        ("previous_bookings_not_canceled", "min / max", (P["previous_bookings_not_canceled"][6],
                                                         P["previous_bookings_not_canceled"][9]), (0, 72)),
    ]
    for col, what, got, want in prof_checks:
        rec("Expected values (brief): column profile (A1)", f"{col}: {what}", "", "", " / ".join(map(fmt, got)),
            " / ".join(map(fmt, want)), tuple(got) == tuple(want))
    for col, needle in (("hotel_name", "City Hotel (79,330); Resort Hotel (40,060)"),
                        ("meal", "BB (92,310); HB (14,463); SC (10,650); Undefined (1,169); FB (798)"),
                        ("is_canceled", "No (75,166); Yes (44,224)"),
                        ("reservation_status", "Check-Out (75,166); Canceled (43,017); No-Show (1,207)"),
                        ("babies", "0 (118,473); 1 (900); 2 (15)"),
                        ("required_car_parking_spaces", "0 (111,974); 1 (7,383); 2 (28); 3 (3); 8 (2)"),
                        ("customer_type", "Transient (89,613); Transient-Party (25,124); Contract (4,076); Group (577)"),
                        ("deposit_type", "No Deposit (104,641); Non Refund (14,587); Refundable (162)"),
                        ("is_repeated_guest", "No (115,580); Yes (3,810)"),
                        ("agent", "NULL (16,340)"), ("company", "NULL (112,593)"),
                        ("children", "0 (110,796); 1 (4,861); 2 (3,652); 3 (76); NA (4)"),
                        ):
        top = str(P[col][11])
        rec("Expected values (brief): column profile (A1)", f"{col}: top values", "", "", top, needle, needle in top)
    odd = str(P["country"][13])
    rec("Expected values (brief): column profile (A1)", "country: 178 distinct incl. the text NULL (488 rows)", "", "",
        f"{P['country'][5]} distinct; {odd}", '178; 488 text "NULL"', P["country"][5] == 178 and odd.startswith(
            '488 text "NULL"'))
    rec("Expected values (brief): column profile (A1)", "29 columns profiled", "", "", len(P), 29, len(P) == 29)

    # ---- Checks sheet and error scan
    for num, lab, left, right, res_ in X["checks"]:
        rec("Checks sheet", f"#{num} {lab}", "", "", f"{fmt(left)} vs {fmt(right)} → {res_}", "TRUE", res_ is True)
    rec("Checks sheet", "Banner A3", "", "", X["banner"], "ALL CHECKS PASS", X["banner"] == "ALL CHECKS PASS")
    rec("Checks sheet", "tblClean: error values in derived columns (recalculated)", "", "", X["clean_errors"], 0,
        X["clean_errors"] == 0)
    rec("Checks sheet", "tblClean: derived cells left without a calculated value", "", "", X["clean_none"], 0,
        X["clean_none"] == 0)

    # ---- Sensitivity Part 2 brief comparison (reported; the known brief slip is shown, not hidden)
    s2_report = []
    for blk, want in EXP_LARGEST.items():
        x = X[("s2", blk)]
        s2_report.append((blk, Cv[("s2", blk, "largest")], Cv[("s2", blk, "where")], want, x["match"], x["why"]))

    # ---- structure and integrity
    print("Structural checks ...", flush=True)
    structure(src, built, recalc, rec)

    # ---- the Word file
    viewed_tokens = set(re.findall(r"\d+", VIEWED))
    tokens, kinds, unmatched, text = check_docx_numbers(docx_path, X["shown_text"], viewed_tokens)
    rec("Word file", f"Every number in the notes appears in the workbook's verified values ({len(tokens)} numbers: "
        + ", ".join(f"{v} {k}" for k, v in kinds.items()) + ")", "", "", ", ".join(unmatched) or "none unmatched",
        "none unmatched", not unmatched)
    rec("Word file", "No em dashes", "", "", text.count("—"), 0, "—" not in text)
    n_pages, broken, outside = pdf_checks(pdf_path, text)
    rec("Word file (rendered PDF)", f"No word broken mid-word ({n_pages} pages checked)", "", "",
        ", ".join(broken) or "none", "none", not broken)
    rec("Word file (rendered PDF)", "No text outside the 2.54 cm side margins (tables fit the page)", "", "",
        ", ".join(outside) or "none", "none", not outside)
    for word in ("canceled", "canceling", "behavior", "color", "analyze", "organization"):
        # case-sensitive with word boundaries: the data value "Canceled" and field names such as is_canceled
        # (underscore = word character) are not counted; ordinary prose must use Australian spelling
        hits = re.findall(rf"\b{word}\b", text)
        rec("Word file", f"US spelling '{word}' in prose", "", "", len(hits), 0, not hits)
    from docx import Document
    d = Document(docx_path)
    sec = d.sections[0]
    cm = lambda emu: round(emu / 360000, 2)
    rec("Word file", "A4 page, 2.54 cm margins", "", "",
        f"{cm(sec.page_width)} × {cm(sec.page_height)}; {cm(sec.left_margin)} / {cm(sec.right_margin)} / "
        f"{cm(sec.top_margin)} / {cm(sec.bottom_margin)}", "21.0 × 29.7; 2.54 all",
        (cm(sec.page_width), cm(sec.page_height)) == (21.0, 29.7) and
        all(cm(m) == 2.54 for m in (sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin)))
    st = d.styles["Normal"]
    rec("Word file", "Body: Times New Roman 12 pt, 1.5 line spacing, justified", "", "",
        f"{st.font.name} {st.font.size.pt} pt; spacing rule {st.paragraph_format.line_spacing_rule}; "
        f"{st.paragraph_format.alignment}", "Times New Roman 12 pt; ONE_POINT_FIVE; JUSTIFY",
        st.font.name == "Times New Roman" and st.font.size.pt == 12 and
        str(st.paragraph_format.line_spacing_rule).startswith("ONE_POINT_FIVE") and
        str(st.paragraph_format.alignment).startswith("JUSTIFY"))
    heads = [t.rows[0] for t in d.tables]
    rep_hdr = all(r._tr.trPr is not None and r._tr.trPr.find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblHeader") is not None for r in heads)
    rec("Word file", "Every table's header row repeats on page breaks", "", "", rep_hdr, True, rep_hdr)
    t1 = d.tables[0]
    rec("Word file", "Issue table: 20 rows × 6 columns", "", "", f"{len(t1.rows) - 1} × {len(t1.columns)}", "20 × 6",
        len(t1.rows) - 1 == 20 and len(t1.columns) == 6)
    for want in ("Data quality notes: BUSM1042 Assignment 3 hotel bookings dataset",
                 "Prepared for the group report; every count is linked to BUSM1042_A3_Group_Clean_Data.xlsx",
                 "Table X: Data quality checks and treatment",
                 "Source: Group analysis of the supplied hotel bookings dataset (n = 119,390); quality dimensions "
                 "from Vidgen, Kirshner and Tan (2019, pp. 57–58).",
                 "Antonio, N, de Almeida, A & Nunes, L 2019, 'Hotel booking demand datasets', Data in Brief, vol. 22, "
                 "pp. 41–49, viewed", "Vidgen, R, Kirshner, S & Tan, F 2019, Business analytics: a management "
                 "approach, Red Globe Press, London."):
        rec("Word file", f"Contains: {want[:70]}…", "", "", want in text, True, want in text)
    words1 = len(d.paragraphs[3].text.split())
    rec("Word file", "Section 1 length (about 120 words)", "", "", words1, "about 120", 100 <= words1 <= 140)

    # ---- report
    n_fail = sum(1 for r in results if r[-1] == "FAIL")
    n_slip = sum(1 for r in results if r[-1] == "DIFFERS FROM BRIEF")
    out = ["# Verification report: group clean dataset and data-quality notes", "",
           f"Generated by `verify_group.py` on {datetime.now():%d %B %Y %H:%M}.", "",
           f"**Overall: {n_fail} FAIL** ({len(results) - n_fail - n_slip} of "
           f"{len(results)} checks PASS, {n_fail} FAIL, {n_slip} DIFFERS FROM BRIEF). The workbook's Checks banner reads "
           f"{X['banner']}.", ""] + ([f"**{n_slip} brief value differs from the data by rounding only** "
                                       "(is_repeated_guest largest change: brief 0.18, exact 0.1746 in all three "
                                       "methods). It is shown, not forced; see Deviations.", ""] if n_slip else []) + [
           "- Method A = pandas on the original file; Method B = a plain Python row loop over the original file "
           "(openpyxl read-only, no pandas); Method C = the workbook's Excel formulas after `soffice --headless "
           "--calc --convert-to xlsx`, read with openpyxl `data_only=True`.",
           "- Counts must agree exactly and rates within 1e-12. Expected values from the brief: counts exactly; rates "
           "given to 1 dp are compared after Decimal half-up rounding to 1 dp; rates given to 2 dp must lie within "
           "±0.005 percentage points; point changes are compared after rounding to 2 dp.", "",
           "## Build facts", "",
           f"- Workbook: {os.path.getsize(built) / 1e6:.1f} MB; LibreOffice load + full recalculation + save: "
           f"{open(os.path.join(os.path.dirname(recalc), 'recalc_time.txt')).read().strip()} s",
           f"- Word file: {os.path.getsize(docx_path) / 1e3:.0f} KB",
           f"- Verification run time: {time.time() - t0:.0f} s", "",
           "## The 20 data-quality issues", "",
           "| ID | Dimension | What was found | Rows affected (workbook) | Expected | Match? |", "|---|---|---|---|---|---|"]
    for did in EXP_DQ:
        row = X[("dq", did)]["row"]
        esc = lambda s: str(s).replace("|", "\\|").replace("\n", " ")
        rows_v = f"{row[4]:,}" if isinstance(row[4], int) else row[4]
        exp_v = f"{row[5]:,}" if isinstance(row[5], int) else row[5]
        out.append(f"| {did} | {esc(row[1])} | {esc(row[3])} | {rows_v} | {exp_v} | {row[6]} |")
    out += ["", "## Largest change per one-variable table vs the brief (Sensitivity Part 2)", "",
            "| Table | Largest change (points, exact) | Category | Brief | Match? |", "|---|---|---|---|---|"]
    for blk, v, where, want, match, why in s2_report:
        out.append(f"| {blk} | {v:.4f} | {where} | {want:.2f} | {match} |")
    section = None
    for sec_, name, a, b, c, exp, st_ in results:
        if sec_ != section:
            out += ["", f"## {sec_}", "", "| Check | Method A (pandas) | Method B (row loop) | Method C (workbook) | "
                                         "Expected | Result |", "|---|---|---|---|---|---|"]
            section = sec_
        esc = lambda s: str(s).replace("|", "\\|").replace("\n", " ⏎ ")
        out.append(f"| {esc(name)} | {esc(a)} | {esc(b)} | {esc(c)} | {esc(exp)} | {st_} |")
    out += ["", "## Deviations from the brief", ""]
    out += DEVIATIONS
    with open(report, "w") as fh:
        fh.write("\n".join(out) + "\n")
    print(f"{len(results) - n_fail - n_slip} / {len(results)} PASS, {n_fail} FAIL, {n_slip} DIFFERS FROM BRIEF; "
          f"report written to {report}")
    for r in results:
        if r[-1] == "FAIL":
            print("FAIL:", r)
    return n_fail


VIEWED = "3 November 2026"
BRIEF_SLIPS = {"is_repeated_guest"}  # brief value 0.18; all three methods give 0.1746 (rounds to 0.17)
DEVIATIONS = [
    "- **Brief value for is_repeated_guest (Sensitivity Part 2):** the brief lists 0.18 points as the largest change. "
    "All three methods give 0.1746 points (Yes: 552 / 3,810 = 14.488% at baseline; 550 / 3,751 = 14.663% after "
    "excluding the 59 flagged repeat-guest bookings), which rounds to 0.17. The brief's 0.18 looks like a rounding "
    "slip (for example, 14.66 minus a truncated 14.48). The typed Expected cell keeps the brief's 0.18 and its Match? "
    "shows ✘ CHECK with the explanation beside it; no value was changed to force a match. This table is not one of "
    "the A7 checks, so the Checks banner does not include it (an information line on the Checks sheet reports it).",
    "- **DQ13 wording:** the brief says 5,400 is \"about 50 times the median\". The median rate is 94.575, so 5,400 "
    "is about 57 times the median; the workbook text is built from the live values and says \"about 57 times\".",
    "- **Excluding no-shows (A5):** the expected 36.03% is reproduced by counting only 'Canceled' as cancelled while "
    "keeping all 119,390 bookings in the denominator (43,017 / 119,390). Dropping the 1,207 no-show rows from the "
    "denominator as well would give 36.40%; the workbook labels the test precisely.",
    "- **Method C command:** as the brief specifies, the workbook was recalculated with `soffice --headless --calc "
    "--convert-to xlsx`; every derived tblClean cell was confirmed to hold a calculated value afterwards.",
    "- **Column_Profile:** distinct counts and Python types come from the Python profile (a distinct count over "
    "119,390 rows is impractical as a formula); all other profile numbers are live formulas on Raw_Data.",
    "- **Reserved room type order:** shown in letter order (A to P), the natural order of the codes; the brief's "
    "\"A … P\" listing suggests the same.",
    "- **Data in Brief paper:** not attached; the open-access copy (ISCTE-IUL repository) was used for Table 1 "
    "(pp. 43–45), the NULL = not applicable statement (p. 45) and the check-in caution (p. 45).",
]


# ================================================================== stage: render the Word file
def stage_render(docx_path, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path], check=True,
                   capture_output=True, timeout=600)
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    pages = int(re.search(r"Pages:\s+(\d+)", subprocess.run(["pdfinfo", pdf], capture_output=True,
                                                            text=True).stdout).group(1))
    subprocess.run(["pdftoppm", "-png", "-r", "80", pdf, os.path.join(out_dir, "notes_page")], check=True)
    print(f"{pdf}: {pages} pages rendered to PNG")


if __name__ == "__main__":
    stage = sys.argv[1]
    if stage == "recalc":
        stage_recalc(*sys.argv[2:4])
    elif stage == "check":
        sys.exit(1 if stage_check(*sys.argv[2:8]) else 0)
    elif stage == "render":
        stage_render(*sys.argv[2:4])
