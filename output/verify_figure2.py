"""Verify BUSM1042_A3_Figure2_FINAL.xlsx (PROMPT 5, Part B and the Figure 2 parts of Parts D and F).

Usage:  python3 -I verify_figure2.py <source.xlsx> <built.xlsx> <work_dir> <report.md>

Method A = pandas on the ORIGINAL source; Method B = a plain row loop (openpyxl read-only, no pandas); Method C = the
workbook's own formulas after `soffice --headless --calc --convert-to xlsx`, read with data_only=True. Also: expected
values from the brief, the displayed axis labels character for character, the chart XML, a regional-setting test
(Vietnamese and German number formats) and a 220 dpi render of the Figure2 sheet.
"""
import json
import os
import re
import subprocess
import sys
import zipfile
from datetime import datetime

SRC, BUILT, WORK, REPORT = sys.argv[1:5]
BANDS = [("0–7", 0, 7), ("8–30", 8, 30), ("31–90", 31, 90), ("91–180", 91, 180), ("181–365", 181, 365),
         ("Over 365", 366, 100000)]
EXP_LABELS = ["0–7 days\n1,902 cancelled", "8–30 days\n5,283 cancelled", "31–90 days\n11,141 cancelled",
              "91–180 days\n11,821 cancelled", "181–365 days\n11,947 cancelled",
              "Over 365 days\n2,130 cancelled\n59% non-refundable"]
EXP_ALL = [9.6, 27.9, 37.7, 44.7, 55.5, 67.7]
EXP_ND = [9.2, 26.0, 32.2, 34.9, 39.8, 20.6]
EXP_N = [19746, 18960, 29553, 26439, 21544, 3148]
EXP_C = [1902, 5283, 11141, 11821, 11947, 2130]
EXP_LEGEND = ["All bookings", "No-deposit bookings only (88% of all bookings)"]
EXP_TITLE = "The further ahead tourists book, the more likely they are to cancel"
EXP_CAPTION = ("Figure 2: Cancellation rate by booking lead time, all bookings and no-deposit bookings, July 2022 to "
               "August 2024")
EXP_NOTE = ("Note: Bars show the share of bookings in each lead-time band that were cancelled, including no-shows; "
            "labels under the bands give cancelled bookings. No-deposit bookings are 87.6% of all bookings. The "
            "over-365-day band is shaded lighter because it holds only 2.6% of bookings, 59.2% of them non-refundable "
            "(see Section 6.4).")
EXP_SOURCE = ("Source: Group analysis of the supplied hotel bookings dataset (n = 119,390), adapted from Antonio, de "
              "Almeida and Nunes (2019c).")
results = []


def rec(section, name, got, exp, ok):
    results.append((section, name, str(got).replace("\n", "⏎"), str(exp).replace("\n", "⏎"), "PASS" if ok else "FAIL"))


def band_of(lt):
    return next(b for b, lo, hi in BANDS if lo <= lt <= hi)


# ---------------------------------------------------------------- Method A: pandas
import pandas as pd  # noqa: E402

d = pd.read_excel(SRC, sheet_name="hotels", engine="openpyxl", keep_default_na=False, dtype=object)
d["band"] = d["lead_time"].astype(int).map(band_of)
d["c"] = (d["is_canceled"] == "Yes").astype(int)
nd = d["deposit_type"] == "No Deposit"
A = {}
for b, *_ in BANDS:
    m = d["band"] == b
    A[b] = dict(n=int(m.sum()), c=int(d["c"][m].sum()), nd=int((m & nd).sum()), ndc=int(d["c"][m & nd].sum()),
                ota=int((m & nd & (d["market_segment"] == "Online TA")).sum()),
                otac=int(d["c"][m & nd & (d["market_segment"] == "Online TA")].sum()),
                nr=int((m & (d["deposit_type"] == "Non Refund")).sum()))
A_tot = dict(n=len(d), c=int(d["c"].sum()), nd=int(nd.sum()))

# ---------------------------------------------------------------- Method B: plain row loop
from openpyxl import load_workbook  # noqa: E402

wb = load_workbook(SRC, read_only=True)
it = wb["hotels"].iter_rows(values_only=True)
h = next(it)
ix = {k: i for i, k in enumerate(h)}
B = {b: dict(n=0, c=0, nd=0, ndc=0, ota=0, otac=0, nr=0) for b, *_ in BANDS}
B_tot = dict(n=0, c=0, nd=0)
for r in it:
    b = band_of(r[ix["lead_time"]])
    canc = 1 if r[ix["is_canceled"]] == "Yes" else 0
    dep = r[ix["deposit_type"]]
    k = B[b]
    k["n"] += 1
    k["c"] += canc
    B_tot["n"] += 1
    B_tot["c"] += canc
    if dep == "No Deposit":
        k["nd"] += 1
        k["ndc"] += canc
        B_tot["nd"] += 1
        if r[ix["market_segment"]] == "Online TA":
            k["ota"] += 1
            k["otac"] += canc
    if dep == "Non Refund":
        k["nr"] += 1
wb.close()

# ---------------------------------------------------------------- Method C: the workbook, recalculated by LibreOffice
os.makedirs(WORK, exist_ok=True)


def recalc(locale=None):
    out = os.path.join(WORK, f"recalc_{locale or 'default'}")
    os.makedirs(out, exist_ok=True)
    env = dict(os.environ)
    prof = os.path.abspath(os.path.join(WORK, f"lo_profile_{locale or 'default'}"))
    if locale:
        env.update(LC_ALL=locale, LANG=locale)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to",
                    "xlsx", "--outdir", out, BUILT], capture_output=True, timeout=900, env=env)
    return os.path.join(out, os.path.basename(BUILT))


def read_c(path):
    w = load_workbook(path, data_only=True)
    q, f, ck = w["Q2_Summary"], w["Figure2"], w["Checks"]
    C = {}
    for i, (b, *_) in enumerate(BANDS):
        r = 6 + i
        v = [q.cell(row=r, column=j).value for j in range(1, 17)]
        C[b] = dict(n=v[3], c=v[4], rate=v[5], nd=v[6], ndc=v[7], ndrate=v[8], ota=v[9], otac=v[10], nr=v[12],
                    nrsh=v[13], shb=v[14], shc=v[15])
    tot = [q.cell(row=12, column=j).value for j in range(1, 17)]
    D = [[q.cell(row=r, column=j).value for j in range(1, 5)] for r in range(50, 58)]
    checks = [(ck.cell(row=r, column=2).value, ck.cell(row=r, column=3).value) for r in range(6, 60)
              if ck.cell(row=r, column=2).value]
    out = dict(a49=q["A49"].value, C=C, tot=tot, D=D, checks=checks, banner=ck["A3"].value, caption=f["A1"].value, note=f["A24"].value,
               source=f["A25"].value, print_area=f.print_area)
    w.close()
    return out


R = read_c(recalc())
C = R["C"]

# ---------------------------------------------------------------- 1. three-way agreement and expected values
for i, (b, *_) in enumerate(BANDS):
    for k in ("n", "c", "nd", "ndc", "ota", "otac", "nr"):
        a, bb, c = A[b][k], B[b][k], C[b][k]
        rec("Three-way agreement (Table A counts)", f"{b} | {k}", f"A {a} / B {bb} / C {c}", "all equal",
            a == bb == c)
    for name, num, den, key in (("rate (all)", "c", "n", "rate"), ("no-deposit rate", "ndc", "nd", "ndrate"),
                                ("Non Refund share", "nr", "n", "nrsh")):
        a, bb, c = A[b][num] / A[b][den], B[b][num] / B[b][den], C[b][key]
        rec("Three-way agreement (Table A rates)", f"{b} | {name}", f"A {a:.12f} / B {bb:.12f} / C {c:.12f}",
            "within 1e-12", abs(a - bb) < 1e-12 and abs(a - c) < 1e-12)
    rec("Expected values (brief)", f"{b} | bookings / cancelled", f"{C[b]['n']:,} / {C[b]['c']:,}",
        f"{EXP_N[i]:,} / {EXP_C[i]:,}", (C[b]["n"], C[b]["c"]) == (EXP_N[i], EXP_C[i]))
    rec("Expected values (brief)", f"{b} | all / no-deposit rate (1 dp)",
        f"{C[b]['rate'] * 100:.1f}% / {C[b]['ndrate'] * 100:.1f}%", f"{EXP_ALL[i]}% / {EXP_ND[i]}%",
        round(C[b]["rate"] * 100, 1) == EXP_ALL[i] and round(C[b]["ndrate"] * 100, 1) == EXP_ND[i])
tot = R["tot"]
rec("Expected values (brief)", "Totals: bookings / cancelled / no-deposit", f"{tot[3]:,} / {tot[4]:,} / {tot[6]:,}",
    "119,390 / 44,224 / 104,641", (tot[3], tot[4], tot[6]) == (119390, 44224, 104641) ==
    (A_tot["n"], A_tot["c"], A_tot["nd"]) == (B_tot["n"], B_tot["c"], B_tot["nd"]))
over = C["Over 365"]
rec("Expected values (brief)", "Over 365: share of bookings / non-refundable share (N11)",
    f"{over['shb'] * 100:.1f}% / {over['nrsh'] * 100:.1f}%", "2.6% / 59.2%",
    (round(over["shb"] * 100, 1), round(over["nrsh"] * 100, 1)) == (2.6, 59.2))
rec("Expected values (brief)", "No-deposit share of all bookings", f"{tot[6] / tot[3] * 100:.1f}%", "87.6%",
    round(tot[6] / tot[3] * 100, 1) == 87.6)
share = sum(C[b]["c"] for b in ("31–90", "91–180", "181–365"))
rec("Expected values (brief)", "31–365 days: cancellations and share", f"{share:,} of {tot[4]:,} = "
    f"{share / tot[4] * 100:.1f}%", "34,909 of 44,224 = 78.9%", share == 34909 and round(share / tot[4] * 100, 1) == 78.9)

# ---------------------------------------------------------------- 2. Table D and Figure2 sheet text, character for character
D = R["D"]
rec("Table D (Q2_Summary A49:D57)", "A49 title", R["a49"], "Table D: Chart data (feeds Figure 2)",
    R["a49"] == "Table D: Chart data (feeds Figure 2)")
rec("Table D (Q2_Summary A49:D57)", "Row 50 headers (legend entries)", " | ".join(map(str, D[0])),
    "Category label (axis) | All bookings | No-deposit bookings only (88% of all bookings) | Check: equals Table A?",
    D[0] == ["Category label (axis)"] + EXP_LEGEND + ["Check: equals Table A?"])
for i in range(6):
    row = D[1 + i]
    rec("Table D (Q2_Summary A49:D57)", f"A{51 + i} displayed label (exact characters)", row[0], EXP_LABELS[i],
        row[0] == EXP_LABELS[i])
    rec("Table D (Q2_Summary A49:D57)", f"B{51 + i}:C{51 + i} rates (1 dp) and D{51 + i} check",
        f"{row[1] * 100:.1f}% / {row[2] * 100:.1f}% / {row[3]}", f"{EXP_ALL[i]}% / {EXP_ND[i]}% / True",
        round(row[1] * 100, 1) == EXP_ALL[i] and round(row[2] * 100, 1) == EXP_ND[i] and row[3] is True)
rec("Table D (Q2_Summary A49:D57)", "A57 note", D[7][0],
    "Each bar is the share of bookings in that band that were cancelled (is_canceled = Yes, including no-shows). "
    "Source: Table A above.", D[7][0] == "Each bar is the share of bookings in that band that were cancelled "
    "(is_canceled = Yes, including no-shows). Source: Table A above.")
for name, got, exp in (("A1 caption", R["caption"], EXP_CAPTION), ("Note below the chart", R["note"], EXP_NOTE),
                       ("Source below the chart", R["source"], EXP_SOURCE)):
    rec("Figure2 sheet text", name, got, exp, got == exp)
rec("Figure2 sheet text", "Print area", R["print_area"], "A1:J28", str(R["print_area"]).replace("$", "").endswith(
    "A1:J28"))
for lab, res in R["checks"]:
    rec("Checks sheet", lab, res, True, res is True)
rec("Checks sheet", "Banner A3", R["banner"], "ALL CHECKS PASS", R["banner"] == "ALL CHECKS PASS")

# ---------------------------------------------------------------- 3. regional-setting test
for loc in ("vi_VN.UTF-8", "de_DE.UTF-8"):
    L = read_c(recalc(loc))
    same = ([r[0] for r in L["D"][1:7]] == EXP_LABELS and L["D"][0][2] == EXP_LEGEND[1] and L["note"] == EXP_NOTE
            and L["source"] == EXP_SOURCE)
    rec("Regional setting", f"{loc}: axis labels, legend, note and source identical to English; banner",
        f"{'identical' if same else 'DIFFERENT'}; {L['banner']}", "identical; ALL CHECKS PASS",
        same and L["banner"] == "ALL CHECKS PASS")

# ---------------------------------------------------------------- 4. chart XML assertions
z = zipfile.ZipFile(BUILT)
charts = sorted(n for n in z.namelist() if re.match(r"xl/charts/chart\d+\.xml$", n))
rec("Chart XML", "Exactly one chart in the workbook", len(charts), 1, len(charts) == 1)
x = re.sub(r"<(/?)[A-Za-z0-9]+:", r"<\1", z.read(charts[0]).decode("utf8"))
x = re.sub(r' xmlns(:\w+)?="[^"]*"', "", x)
wbx = re.sub(r"<(/?)[A-Za-z0-9]+:", r"<\1", z.read("xl/workbook.xml").decode("utf8"))
sheets = re.findall(r'<sheet [^>]*name="([^"]+)"', wbx)
rels = z.read(f"xl/worksheets/_rels/sheet{sheets.index('Figure2') + 1}.xml.rels").decode("utf8")
drw = re.sub(r"<(/?)[A-Za-z0-9]+:", r"<\1", z.read("xl/drawings/drawing1.xml").decode("utf8"))


def blk(tag, s=x):
    m = re.search(rf"<{tag}>(.*?)</{tag}>", s, flags=re.S)
    return m.group(1) if m else ""


sers = re.findall(r"<ser>(.*?)</ser>", x, flags=re.S)
val, cat = blk("valAx"), blk("catAx")
ext = re.search(r'<ext cx="(\d+)" cy="(\d+)"', drw)
frm = re.search(r"<from><col>(\d+)</col><colOff>\d+</colOff><row>(\d+)</row>", drw)


def ser_info(s):
    dpts = re.findall(r"<dPt>(.*?)</dPt>", s, flags=re.S)
    dlbls = re.findall(r"<dLbl>(.*?)</dLbl>", s, flags=re.S)
    return dict(
        tx=re.search(r"<tx><strRef><f>(.*?)</f>", s).group(1).replace("$", ""), val=re.search(r"<val><numRef><f>(.*?)</f>", s).group(1),
        cat=re.search(r"<cat><(?:str|num)Ref><f>(.*?)</f>", s).group(1),
        fill=re.search(r'<order val="\d+"/>.*?<spPr><solidFill><srgbClr val="(\w+)"', s, flags=re.S).group(1),
        dpt=[(re.search(r'<idx val="(\d+)"', p).group(1), re.search(r'<srgbClr val="(\w+)"', p).group(1)) for p in dpts],
        dlbl=[(re.search(r'<idx val="(\d+)"', p).group(1), re.search(r'<srgbClr val="(\w+)"', p).group(1),
               re.search(r'<showVal val="(\d)"', p).group(1), re.search(r'b="(\d)"', p).group(1)) for p in dlbls],
        lblpos=re.findall(r'<dLblPos val="(\w+)"', s), fmt=re.findall(r'<numFmt formatCode="([^"]+)"', s),
        sz=re.findall(r'<defRPr sz="(\d+)"', blk("dLbls", s)),
        bold=re.search(r'</dLbl>.*?<defRPr sz="\d+" b="(\d)"', s, flags=re.S))


S = [ser_info(s) for s in sers]
A_ = {
    "barChart, barDir col, grouping clustered": "<barChart>" in x and '<barDir val="col"/>' in x and
                                                '<grouping val="clustered"/>' in x,
    "exactly 2 series; no lineChart or other chart types": len(sers) == 2 and not re.search(
        r"<(line|area|pie|doughnut|scatter|radar|bubble)Chart", x),
    "gapWidth 40, overlap -12": '<gapWidth val="40"/>' in x and '<overlap val="-12"/>' in x,
    "no majorGridlines": "majorGridlines" not in x,
    "value axis deleted, min 0, max 0.8, major unit 0.2": '<delete val="1"/>' in val and '<min val="0"/>' in val and
                                                          '<max val="0.8"/>' in val and '<majorUnit val="0.2"/>' in val,
    "category axis shown, rot 0, tickLblSkip 1, tickMarkSkip 1": '<delete val="0"/>' in cat and 'rot="0"' in cat and
        '<tickLblSkip val="1"/>' in cat and '<tickMarkSkip val="1"/>' in cat,
    "category labels Calibri 8 pt #595959": re.search(r'<txPr><bodyPr rot="0"/><p><pPr><defRPr sz="800" b="0">'
                                                      r'<solidFill><srgbClr val="595959"', cat) is not None,
    "axis line #BFBFBF 0.75 pt": '<ln w="9525"><solidFill><srgbClr val="BFBFBF"/>' in cat,
    "axis title 'Lead time (days between booking and arrival)', 9 pt regular #595959":
        re.search(r'<defRPr sz="900" b="0"><solidFill><srgbClr val="595959"/>.*?<t>Lead time \(days between booking '
                  r'and arrival\)</t>', cat, flags=re.S) is not None,
    "series 1: name B50, values B51:B56, categories A51:A56": (S[0]["tx"], S[0]["val"], S[0]["cat"]) == (
        "'Q2_Summary'!B50", "'Q2_Summary'!$B$51:$B$56", "'Q2_Summary'!$A$51:$A$56"),
    "series 2: name C50, values C51:C56, categories A51:A56": (S[1]["tx"], S[1]["val"], S[1]["cat"]) == (
        "'Q2_Summary'!C50", "'Q2_Summary'!$C$51:$C$56", "'Q2_Summary'!$A$51:$A$56"),
    "series fills #1F5AA6 / #A6A6A6, no outlines": (S[0]["fill"], S[1]["fill"]) == ("1F5AA6", "A6A6A6") and
                                                    x.count("<ln><noFill/>") >= 4,
    "exactly one dPt per series, idx 5, fills #BDD0EA / #DCDCDC": (S[0]["dpt"], S[1]["dpt"]) == (
        [("5", "BDD0EA")], [("5", "DCDCDC")]),
    "one dLbl idx 5 per series: showVal 1, regular, #5B7DB8 / #7F7F7F": (S[0]["dlbl"], S[1]["dlbl"]) == (
        [("5", "5B7DB8", "1", "0")], [("5", "7F7F7F", "1", "0")]),
    "series labels: bold #1F5AA6 (1) and regular #7F7F7F (2)":
        re.search(r'</dLbl>.*?<defRPr sz="850" b="1"><solidFill><srgbClr val="1F5AA6"', sers[0], flags=re.S) is not None
        and re.search(r'</dLbl>.*?<defRPr sz="850" b="0"><solidFill><srgbClr val="7F7F7F"', sers[1], flags=re.S)
        is not None,
    "data labels outEnd, 0.0%, 8.5 pt, on all bars": all(set(s["lblpos"]) == {"outEnd"} and set(s["fmt"]) == {"0.0%"}
                                                          and set(s["sz"]) == {"850"} for s in S),
    "title text exactly, Calibri 12 pt bold #262626, top-left (x 0.01, y 0.01)":
        f"<t>{EXP_TITLE}</t>" in blk("title") and 'sz="1200" b="1"' in blk("title") and "262626" in blk("title") and
        '<x val="0.01"/>' in blk("title") and '<y val="0.01"/>' in blk("title"),
    "legend at the top, Calibri 9 pt #404040": '<legendPos val="t"/>' in x and re.search(
        r'<legend>.*?<defRPr sz="900" b="0"><solidFill><srgbClr val="404040"', x, flags=re.S) is not None,
    "chart area white, no border": "<spPr><solidFill><srgbClr val=\"FFFFFF\"/></solidFill><ln><noFill/>" in x,
    "Calibri throughout; no 3-D": 'typeface="Calibri"' in x and "3D" not in x,
    "size 15.9 × 10.0 cm": abs(int(ext.group(1)) / 360000 - 15.9) < 0.01 and abs(int(ext.group(2)) / 360000 - 10.0)
                           < 0.01,
    "anchored at Figure2!A4": frm.groups() == ("0", "3") and "drawing" in rels,
}
for k, v in A_.items():
    rec("Chart XML", k, v, True, v)

# ---------------------------------------------------------------- 5. render the Figure2 sheet (220 dpi) + chart crop
rd = os.path.join(WORK, "render")
os.makedirs(rd, exist_ok=True)
default_recalc = os.path.join(WORK, "recalc_default", os.path.basename(BUILT))
subprocess.run(["soffice", f"-env:UserInstallation=file://{os.path.abspath(os.path.join(WORK, 'lo_profile_default'))}",
                "--headless", "--convert-to", "pdf", "--outdir", rd, default_recalc], capture_output=True, timeout=900)
pdf = os.path.join(rd, os.path.splitext(os.path.basename(BUILT))[0] + ".pdf")
subprocess.run(["pdftoppm", "-png", "-r", "220", "-f", "1", "-l", "1", "-singlefile", pdf,
                os.path.join(rd, "Figure2_FINAL_render")], check=True)
subprocess.run(["pdftoppm", "-png", "-r", "300", "-f", "1", "-l", "1", "-singlefile", pdf,
                os.path.join(rd, "page300")], check=True)
from PIL import Image  # noqa: E402

im = Image.open(os.path.join(rd, "page300.png")).convert("RGB")
txt = subprocess.run(["pdftotext", "-f", "1", "-l", "1", "-bbox", pdf, "-"], capture_output=True, text=True).stdout
words = re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', txt)
pw = float(re.search(r'<page width="([\d.]+)"', txt).group(1))
k = im.width / pw
y_top = min(float(w[1]) for w in words if w[4] == "The")  # chart title
y_bot = max(float(w[3]) for w in words if w[4] in ("arrival)",))  # x-axis title
x_left = min(float(w[0]) for w in words if float(w[1]) >= y_top - 1 and float(w[3]) <= y_bot + 1)
x_right = max(float(w[2]) for w in words if float(w[1]) >= y_top - 1 and float(w[3]) <= y_bot + 1)
pad = 14
crop = im.crop((int((x_left - pad) * k), int((y_top - pad) * k), int((x_right + pad) * k), int((y_bot + pad) * k)))
crop.save(os.path.join(rd, "Figure2_FINAL_chart_300dpi.png"), dpi=(300, 300))
rec("Render", "Figure2 sheet rendered at 220 dpi; chart cropped at 300 dpi", f"{im.width}×{im.height} page; chart "
    f"{crop.width}×{crop.height} px", "files written", True)

# ---------------------------------------------------------------- report
n_fail = sum(r[-1] == "FAIL" for r in results)
out = ["# Verification report: Figure 2 FINAL", "", f"Generated by `verify_figure2.py` on {datetime.now():%d %B %Y %H:%M}.",
       "", f"**Overall: {'ALL PASS' if not n_fail else f'{n_fail} FAIL'}** ({len(results) - n_fail} of {len(results)} "
       "checks passed).", "",
       "- Method A = pandas on the original file; Method B = a plain row loop over the original file (no pandas); "
       "Method C = the workbook's formulas after `soffice --headless --calc --convert-to xlsx`, read with "
       "`data_only=True`. Counts must match exactly, rates within 1e-12.",
       "- Rendered previews (LibreOffice; checking only): `Figure2_FINAL_render.png` (220 dpi, whole Figure2 sheet) "
       "and `Figure2_FINAL_chart_300dpi.png` (the chart alone). For the report, paste the native chart from Excel as "
       "an Enhanced Metafile.", ""]
sec = None
for s, name, got, exp, st in results:
    if s != sec:
        out += ["", f"## {s}", "", "| Check | Result in the workbook | Expected | Status |", "|---|---|---|---|"]
        sec = s
    e = lambda v: str(v).replace("|", "\\|")
    out.append(f"| {e(name)} | {e(got)} | {e(exp)} | {st} |")
out += ["", "## Notes", "",
        "- Figure 2 only: the cleaning, flags and PivotTables are in the group workbook and were not rebuilt. Table A's "
        "counts are computed from the original source by `build_figure2.py` (and asserted against the brief before "
        "saving); every rate, share, label and note is an Excel formula on them.",
        "- The chart is `build_chart(Q2_Summary, 50, 51, 56, 1, 2, 3)` from `figure2_FINAL_reference.py`, unchanged in "
        "style. Two robustness additions only: the categories are written as a text reference (strRef), and the chart "
        "area is explicitly white.",
        "- Labels and notes avoid TEXT(…,\"#,##0\") and TEXT(…,\"0.0%\"), which follow the computer's regional setting; "
        "they are built so they read the same everywhere (regional-setting test above). The bar value labels use the "
        "number format 0.0%, which Excel shows with the computer's decimal sign: copy the chart with English number "
        "formats so they read 9.6%, not 9,6%.",
        "- LibreOffice renders with Carlito (Calibri metrics). Excel's own rendering may differ by a pixel or two in "
        "label placement; the content is the same.", ""]
with open(REPORT, "w") as fh:
    fh.write("\n".join(out))
print(f"{len(results) - n_fail} / {len(results)} PASS")
for r in results:
    if r[-1] == "FAIL":
        print("FAIL:", r)
sys.exit(1 if n_fail else 0)
