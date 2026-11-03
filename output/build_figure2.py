"""Build BUSM1042_A3_Figure2_FINAL.xlsx: report Figure 2 only (PROMPT 5, Part B), as a native Excel chart.

Usage:  python3 -I build_figure2.py <IBA-A3-Hotel_bookings_dataset-T32026.xlsx> <output.xlsx>

The data cleaning, flags and PivotTables already exist in the group workbook, so this file holds only what the figure
needs: Table A (counts per lead-time band, computed here from the ORIGINAL source file and verified three ways by
verify_figure2.py), Table D (the chart data, all formulas), the Figure2 sheet and a Checks sheet.
The chart comes from figure2_FINAL_reference.py (build_chart and its colour constants), unchanged in style.
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # -I drops the script folder from sys.path
import figure2_FINAL_reference as ref  # noqa: E402

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.chart.data_source import AxDataSource, StrRef  # noqa: E402
from openpyxl.chart.shapes import GraphicalProperties  # noqa: E402
from openpyxl.drawing.line import LineProperties  # noqa: E402
from openpyxl.formatting.rule import FormulaRule  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.worksheet.properties import PageSetupProperties  # noqa: E402

SRC, OUT = sys.argv[1], sys.argv[2]
FONT = "Calibri"
NAVY = "1F3A5F"
F_BODY = Font(name=FONT, size=11)
F_BOLD = Font(name=FONT, size=11, bold=True)
F_HDR = Font(name=FONT, size=11, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
F_TT = Font(name=FONT, size=12, bold=True, color=NAVY)
F_NOTE = Font(name=FONT, size=10, italic=True, color="595959")
F_VAL = Font(name=FONT, size=11, color="0000FF")
THIN = Side(style="thin", color="BFBFBF")
B_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
B_TOT = Border(left=THIN, right=THIN, top=Side(style="medium", color="404040"), bottom=THIN)
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_BAND = PatternFill("solid", fgColor="F2F2F2")
A_HDR = Alignment(horizontal="center", vertical="center", wrap_text=True)
A_NUM = Alignment(horizontal="right", vertical="center")
A_LEFT = Alignment(horizontal="left", vertical="center")
GREEN = (PatternFill("solid", fgColor="C6EFCE"), Font(name=FONT, color="006100", bold=True))
RED = (PatternFill("solid", fgColor="FFC7CE"), Font(name=FONT, color="9C0006", bold=True))
BANDS = [("0–7", 0, 7), ("8–30", 8, 30), ("31–90", 31, 90), ("91–180", 91, 180), ("181–365", 181, 365),
         ("Over 365", 366, 100000)]


# ---------------------------------------------------------------- locale-proof number text for labels and notes
# TEXT(x,"#,##0") and TEXT(x,"0.0%") follow the computer's regional setting (16.340 / 87,6% in Vietnamese or German
# Excel). These builders give the same English text everywhere: TEXT(...,"000") and whole numbers carry no separator.
def thousands(x):
    return f'IF({x}>=1000,INT({x}/1000)&","&TEXT(MOD({x},1000),"000"),INT({x})&"")'


def pct1(x):
    return f'INT(ROUND({x}*1000,0)/10)&"."&MOD(ROUND({x}*1000,0),10)&"%"'


def pct0(x):
    return f'ROUND({x}*100,0)&"%"'


# ---------------------------------------------------------------- counts from the ORIGINAL source (read-only)
t0 = time.time()
print("Reading the source ...", flush=True)
src = load_workbook(SRC, read_only=True)
rows = src["hotels"].iter_rows(values_only=True)
hdr = list(next(rows))
ix = {h: i for i, h in enumerate(hdr)}
counts = {b[0]: dict(n=0, c=0, nd=0, ndc=0, ota=0, otac=0, nr=0) for b in BANDS}
total = 0
for r in rows:
    lt = r[ix["lead_time"]]
    band = next(b for b, lo, hi in BANDS if lo <= lt <= hi)
    canc = r[ix["is_canceled"]] == "Yes"
    dep = r[ix["deposit_type"]]
    k = counts[band]
    k["n"] += 1
    k["c"] += canc
    if dep == "No Deposit":
        k["nd"] += 1
        k["ndc"] += canc
        if r[ix["market_segment"]] == "Online TA":
            k["ota"] += 1
            k["otac"] += canc
    if dep == "Non Refund":
        k["nr"] += 1
    total += 1
src.close()
assert total == 119390, total
EXPECTED_N = [19746, 18960, 29553, 26439, 21544, 3148]
EXPECTED_C = [1902, 5283, 11141, 11821, 11947, 2130]
assert [counts[b[0]]["n"] for b in BANDS] == EXPECTED_N, "band bookings differ from the brief: STOP"
assert [counts[b[0]]["c"] for b in BANDS] == EXPECTED_C, "band cancellations differ from the brief: STOP"

wb = Workbook()
wb.calculation.fullCalcOnLoad = True
fig = wb.active
fig.title = "Figure2"
q = wb.create_sheet("Q2_Summary")
ck = wb.create_sheet("Checks")
rd = wb.create_sheet("README")


def cell(ws, ref_, value, font=F_BODY, fmt=None, align=None, fill=None, border=None):
    c = ws[ref_]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    if fill:
        c.fill = fill
    if border:
        c.border = border
    return c


def header(ws, row, labels, height=46):
    for j, h in enumerate(labels):
        c = ws.cell(row=row, column=1 + j, value=h)
        c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, A_HDR, B_ALL
    ws.row_dimensions[row].height = height


# ---------------------------------------------------------------- Q2_Summary: Table A (rows 5-12) and Table D (49-57)
cell(q, "A1", "Q2 summary for Figure 2: cancellation by lead-time band", F_TITLE)
cell(q, "A2", "Blue counts were computed by build_figure2.py from the original source file (all 119,390 bookings; "
              "cancellation = is_canceled = Yes, including no-shows) and verified three ways by verify_figure2.py. "
              "Every rate, share, label and note is an Excel formula on these counts.", F_NOTE)
cell(q, "A4", "Table A: Cancellation by lead-time band", F_TT)
header(q, 5, ["Lead time band (days)", "Lower bound (days)", "Upper bound (days)", "Bookings", "Cancelled",
              "Cancellation rate (all)", "No-deposit bookings", "No-deposit cancelled", "No-deposit rate",
              "No-deposit Online TA bookings", "No-deposit Online TA cancelled", "No-deposit Online TA rate",
              "Non Refund bookings", "Non Refund share of band", "Share of all bookings",
              "Share of all cancellations"])
for i, (band, lo, hi) in enumerate(BANDS):
    r = 6 + i
    k = counts[band]
    vals = [band, lo, hi, k["n"], k["c"], f"=E{r}/D{r}", k["nd"], k["ndc"], f"=H{r}/G{r}", k["ota"], k["otac"],
            f"=K{r}/J{r}", k["nr"], f"=M{r}/D{r}", f"=D{r}/$D$12", f"=E{r}/$E$12"]
    fmts = [None, "0", "#,##0", "#,##0", "#,##0", "0.0%", "#,##0", "#,##0", "0.0%", "#,##0", "#,##0", "0.0%", "#,##0",
            "0.0%", "0.0%", "0.0%"]
    for j, (v, f) in enumerate(zip(vals, fmts)):
        c = q.cell(row=r, column=1 + j, value=v)
        c.font = F_VAL if isinstance(v, int) and j >= 3 else F_BODY
        c.number_format = f or "General"
        c.alignment = A_LEFT if j == 0 else A_NUM
        c.border = B_ALL
        if i % 2:
            c.fill = FILL_BAND
tot = ["Total", None, None, "=SUM(D6:D11)", "=SUM(E6:E11)", "=E12/D12", "=SUM(G6:G11)", "=SUM(H6:H11)", "=H12/G12",
       "=SUM(J6:J11)", "=SUM(K6:K11)", "=K12/J12", "=SUM(M6:M11)", "=M12/D12", "=SUM(O6:O11)", "=SUM(P6:P11)"]
for j, v in enumerate(tot):
    c = q.cell(row=12, column=1 + j, value=v)
    c.font, c.border = F_BOLD, B_TOT
    c.number_format = ["General", "0", "0", "#,##0", "#,##0", "0.0%", "#,##0", "#,##0", "0.0%", "#,##0", "#,##0",
                       "0.0%", "#,##0", "0.0%", "0.0%", "0.0%"][j]
    c.alignment = A_LEFT if j == 0 else A_NUM
cell(q, "A13", "Bands are inclusive ('Over 365' = 366 days or more). No-deposit = deposit_type 'No Deposit'; Online "
               "TA = market_segment 'Online TA'. Source: IBA-A3-Hotel_bookings_dataset-T32026.xlsx.", F_NOTE)

cell(q, "A49", "Table D: Chart data (feeds Figure 2)", F_TT)
header(q, 50, ["Category label (axis)", "All bookings", None, "Check: equals Table A?"], height=46)
q["C50"] = f'="No-deposit bookings only ("&{pct0("G12/D12")}&" of all bookings)"'  # legend entry, locale-proof
q["C50"].font, q["C50"].fill, q["C50"].alignment, q["C50"].border = F_HDR, FILL_HDR, A_HDR, B_ALL
for i in range(6):
    r, a = 51 + i, 6 + i
    label = f'=A{a}&" days"&CHAR(10)&{thousands(f"E{a}")}&" cancelled"'
    if i == 5:  # Over 365: add the non-refundable share that explains the no-deposit drop
        label += f'&CHAR(10)&{pct0(f"N{a}")}&" non-refundable"'
    for j, (v, f) in enumerate(((label, None), (f"=F{a}", "0.0%"), (f"=I{a}", "0.0%"),
                                (f"=AND(B{r}=F{a},C{r}=I{a})", None))):
        c = q.cell(row=r, column=1 + j, value=v)
        c.font, c.border = F_BODY, B_ALL
        c.number_format = f or "General"
        c.alignment = Alignment(wrap_text=True, vertical="center") if j == 0 else (
            A_NUM if j < 3 else Alignment(horizontal="center", vertical="center"))
    q.row_dimensions[r].height = 47 if i == 5 else 32
q.conditional_formatting.add("D51:D56", FormulaRule(formula=["D51=TRUE"], fill=GREEN[0], font=GREEN[1]))
q.conditional_formatting.add("D51:D56", FormulaRule(formula=["D51<>TRUE"], fill=RED[0], font=RED[1]))
cell(q, "A57", "Each bar is the share of bookings in that band that were cancelled (is_canceled = Yes, including "
               "no-shows). Source: Table A above.", F_NOTE)
for col, w in zip("ABCDEFGHIJKLMNOP", [24, 10, 10, 11, 11, 12, 12, 12, 11, 13, 13, 12, 11, 12, 11, 12]):
    q.column_dimensions[col].width = w
q.column_dimensions["C"].width = 24
q.sheet_view.showGridLines = False
q.sheet_properties.tabColor = "70AD47"
q.freeze_panes = "B6"

# ---------------------------------------------------------------- Figure2 sheet: the chart, exactly as the reference
cell(fig, "A1", "Figure 2: Cancellation rate by booking lead time, all bookings and no-deposit bookings, July 2022 to "
                "August 2024", Font(name=FONT, size=12, bold=True))
cell(fig, "A2", "(For Word: the caption goes above the pasted chart; the note and source go below it.)",
     Font(name=FONT, size=10, italic=True, color="595959"))
ch = ref.build_chart(q, 50, 51, 56, 1, 2, 3)  # default fade_points=(5,): the Over 365 pair is faded
# Not restyling, only making the file robust in Excel: text categories as a string reference (multi-line labels),
# and an explicit white chart area (the reference relies on Excel's default white).
for s in ch.series:
    s.cat = AxDataSource(strRef=StrRef(f="'Q2_Summary'!$A$51:$A$56"))
ch.graphical_properties = GraphicalProperties(solidFill="FFFFFF", ln=LineProperties(noFill=True))
fig.add_chart(ch, "A4")
NOTE_ROW, SRC_ROW = 24, 25
fig[f"A{NOTE_ROW}"] = (
    '="Note: Bars show the share of bookings in each lead-time band that were cancelled, including no-shows; labels '
    'under the bands give cancelled bookings. No-deposit bookings are "&' + pct1("Q2_Summary!G12/Q2_Summary!D12") +
    '&" of all bookings. The over-365-day band is shaded lighter because it holds only "&' + pct1("Q2_Summary!O11") +
    '&" of bookings, "&' + pct1("Q2_Summary!N11") + '&" of them non-refundable (see Section 6.4)."')
fig[f"A{SRC_ROW}"] = ('="Source: Group analysis of the supplied hotel bookings dataset (n = "&' +
                      thousands("Q2_Summary!D12") + '&"), adapted from Antonio, de Almeida and Nunes (2019c)."')
for col in "ABCDEFGHIJ":
    fig.column_dimensions[col].width = 8.0
for r, h in ((1, 32), (2, 15), (NOTE_ROW, 40), (SRC_ROW, 27)):  # 3-line note, 2-line source
    fig.merge_cells(f"A{r}:J{r}")
    fig[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    fig.row_dimensions[r].height = h
for r in (NOTE_ROW, SRC_ROW):
    fig[f"A{r}"].font = Font(name=FONT, size=9, color="404040")
fig.sheet_view.showGridLines = False
fig.sheet_properties.tabColor = "ED7D31"
fig.page_setup.paperSize = fig.PAPERSIZE_A4
fig.page_setup.orientation = "portrait"
fig.page_setup.fitToWidth = 1
fig.page_setup.fitToHeight = 1
fig.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
fig.print_area = "A1:J28"

# ---------------------------------------------------------------- Checks (Part D items that concern the figure)
cell(ck, "A1", "Checks for Figure 2 (live TRUE/FALSE)", F_TITLE)
header(ck, 5, ["#", "Check", "Result"], height=22)
QS = "Q2_Summary"
CHECKS = [
    ("Table A totals: 119,390 bookings and 44,224 cancelled", f"=AND({QS}!D12=119390,{QS}!E12=44224)"),
    ("Table D series equal Table A in all six bands", f"=COUNTIF({QS}!D51:D56,TRUE)=6"),
] + [(f"Axis label {i + 1} shows that band's cancelled bookings",
      f'=ISNUMBER(SEARCH({thousands(f"{QS}!E{6 + i}")}&" cancelled",{QS}!A{51 + i}))') for i in range(6)] + [
    ("No axis label contains the word 'bookings'", f'=SUMPRODUCT(--ISNUMBER(SEARCH("bookings",{QS}!A51:A56)))=0'),
    ("Over 365 label shows the non-refundable share",
     f'=ISNUMBER(SEARCH({pct0(f"{QS}!N11")}&" non-refundable",{QS}!A56))'),
    ("Chart scale fits (every rate below 80%)", f"=MAX({QS}!B51:C56)<0.8"),
    ("Title claim: the all-bookings rate rises in every band",
     f"=AND({QS}!F6<{QS}!F7,{QS}!F7<{QS}!F8,{QS}!F8<{QS}!F9,{QS}!F9<{QS}!F10,{QS}!F10<{QS}!F11)"),
    ("No-deposit rate rises across the first five bands",
     f"=AND({QS}!I6<{QS}!I7,{QS}!I7<{QS}!I8,{QS}!I8<{QS}!I9,{QS}!I9<{QS}!I10)"),
    ("Most cancellations come from 31–365 days", f"=SUM({QS}!E8:E10)/{QS}!E12>0.5"),
    ("Legend: C50 reads exactly 'No-deposit bookings only (88% of all bookings)'",
     f'={QS}!C50="No-deposit bookings only (88% of all bookings)"'),
    ("Fade reason: Over 365 is under 5% of bookings and over 50% non-refundable",
     f"=AND({QS}!O11<0.05,{QS}!N11>0.5)"),
]
EXP_RATES = [(6, "F", 0.096), (7, "F", 0.279), (8, "F", 0.377), (9, "F", 0.447), (10, "F", 0.555), (11, "F", 0.677),
             (6, "I", 0.092), (7, "I", 0.260), (8, "I", 0.322), (9, "I", 0.349), (10, "I", 0.398), (11, "I", 0.206)]
for r, col, v in EXP_RATES:
    band = BANDS[r - 6][0]
    CHECKS.append((f"{'All' if col == 'F' else 'No-deposit'} rate, {band} days, rounds to {v * 100:.1f}% (brief)",
                   f"=ROUND({QS}!{col}{r},3)={v}"))
for i, (lab, f) in enumerate(CHECKS):
    r = 6 + i
    for j, (v, al) in enumerate(((i + 1, Alignment(horizontal="center")), (lab, A_LEFT),
                                 (f, Alignment(horizontal="center")))):
        c = ck.cell(row=r, column=1 + j, value=v)
        c.font, c.border, c.alignment = F_BODY, B_ALL, al
        if i % 2:
            c.fill = FILL_BAND
END = 5 + len(CHECKS)
rng = f"$C$6:$C${END}"
ck.conditional_formatting.add(f"C6:C{END}", FormulaRule(formula=["C6=TRUE"], fill=GREEN[0], font=GREEN[1]))
ck.conditional_formatting.add(f"C6:C{END}", FormulaRule(formula=["C6<>TRUE"], fill=RED[0], font=RED[1]))
cell(ck, "A3", f'=IF(AND(COUNTIF({rng},FALSE)=0,COUNTIF({rng},TRUE)=ROWS({rng})),"ALL CHECKS PASS","CHECK FAILED")',
     Font(name=FONT, size=20, bold=True))
ck.conditional_formatting.add("A3", FormulaRule(formula=['A3="ALL CHECKS PASS"'],
                                                font=Font(name=FONT, size=20, bold=True, color="006100")))
ck.conditional_formatting.add("A3", FormulaRule(formula=['A3<>"ALL CHECKS PASS"'],
                                                font=Font(name=FONT, size=20, bold=True, color="C00000")))
ck.row_dimensions[3].height = 32
cell(ck, "A4", f'=COUNTIF({rng},TRUE)&" of "&ROWS({rng})&" checks are TRUE."', Font(name=FONT, size=10))
for col, w in zip("ABC", [5, 78, 10]):
    ck.column_dimensions[col].width = w
ck.sheet_view.showGridLines = False
ck.sheet_properties.tabColor = "C00000"

# ---------------------------------------------------------------- README
lines = [
    ("What this is", "Report Figure 2 for business question Q2, 'How does booking lead time affect the likelihood of "
                     "cancellation?', as a native Excel chart, with the data that feeds it. Data cleaning, flags and "
                     "PivotTables are in the group workbook (BUSM1042_A3_Group_Clean_Data.xlsx) and are not repeated."),
    ("Sheets", "Figure2: the chart, caption, note and source. Q2_Summary: Table A (A4) and Table D, the chart data "
               "(A49). Checks: live TRUE/FALSE checks with a banner in A3. README: this page."),
    ("Scenario rule", "Every booking is treated as an Australian tourist's; country is origin, not citizenship, and is "
                      "never used to filter."),
    ("Bands", "0–7, 8–30, 31–90, 91–180, 181–365 and Over 365 (366 or more) days, inclusive."),
    ("Into Word", "Open in desktop Excel, wait for the calculation to finish, then Save. Click the chart → Copy → in "
                  "Word, Paste Special → Picture (Enhanced Metafile). Put the caption above it and the note and source "
                  "below it."),
    ("Regional setting", "Labels and notes are built so they read the same in any regional setting (16,340 not "
                         "16.340). The bar labels use Excel's number format 0.0%, which follows the computer's "
                         "setting: copy the chart with English number formats so they read 9.6%, not 9,6%."),
    ("FIGURE 2 DESIGN", None),
    ("Action title", "States the takeaway, top-left (Knaflic 2015, ch. 8); visuals need descriptive titles and a clear "
                     "objective (Tableau guide)."),
    ("Bars on a zero baseline", "Familiar and read by length (Knaflic 2015, pp. 50–51); Lecture 6 lists omitted "
                                "baselines and truncated scales as mistakes."),
    ("Natural order", "Bands kept in lead-time order (Knaflic 2015, p. 58)."),
    ("Labels instead of a y-axis", "The exact values matter and an axis would only repeat them (Knaflic 2015, p. 52); "
                                   "no gridlines or border (Knaflic 2015, p. 194; Tufte's data-ink idea)."),
    ("One focus colour", "Blue #1F5AA6 for the answer (all bookings), grey #A6A6A6 for the context series (Knaflic "
                         "2015, ch. 8; Duarte 2014); Lecture 6 warns against misleading colour contrast."),
    ("No 3D, pie or secondary axis", "Knaflic (2015, pp. 66–68); Lecture 6. A dual-axis combo was tested and "
                                     "rejected."),
    ("Context under the bars", "Cancelled bookings and the deposit explanation sit under the bars they explain "
                               "(Geckoboard: numbers need context; round where precision adds nothing)."),
    ("Faded, not removed", "The Over 365 pair (2.6% of bookings, 59% non-refundable) is pale blue #BDD0EA and pale "
                           "grey #DCDCDC (Knaflic 2015, p. 195), so the clean 0–365 rise reads first; nothing is "
                           "hidden."),
    ("Subset named in the legend", "'No-deposit bookings only (88% of all bookings)', so the two series cannot be "
                                   "misread as separate groups."),
]
cell(rd, "A1", "README: Figure 2 (FINAL)", F_TITLE)
r = 3
for k, v in lines:
    if v is None:
        cell(rd, f"A{r}", k, F_TT)
    else:
        cell(rd, f"A{r}", k, F_BOLD, align=Alignment(vertical="top", wrap_text=True))
        cell(rd, f"B{r}", v, align=Alignment(vertical="top", wrap_text=True))
        rd.row_dimensions[r].height = 15 * max(1, -(-len(v) // 105))
    r += 1
rd.column_dimensions["A"].width = 28
rd.column_dimensions["B"].width = 100
rd.sheet_view.showGridLines = False
rd.sheet_properties.tabColor = "808080"

wb.active = 0
wb.save(OUT)
print(f"Saved {OUT} in {time.time() - t0:.0f}s")
