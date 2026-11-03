"""Build BUSM1042_A3_Data_Quality_Notes.docx from the RECALCULATED group workbook.

Usage:  python3 -I build_quality_notes.py <recalculated BUSM1042_A3_Group_Clean_Data.xlsx> <output.docx> "<viewed date>"

Every number in the notes is read from the workbook's calculated values (LibreOffice recalculation), so the Word
file and the workbook cannot disagree. Format: A4, 2.54 cm margins, Times New Roman 12 pt, 1.5 line spacing,
justified body, numbered headings, 9 pt tables with a navy header row that repeats across pages.
"""
import sys
from collections import OrderedDict

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import load_workbook

XLSX, OUT, VIEWED = sys.argv[1], sys.argv[2], sys.argv[3]
NAVY = "1F3A5F"
FONT = "Times New Roman"

# ================================================================ read the calculated workbook
wb = load_workbook(XLSX, data_only=True)
dq_ws, se, ov, ck = wb["Data_Quality_Log"], wb["Sensitivity"], wb["OneVar_Summary"], wb["Checks"]
assert ck["A3"].value == "ALL CHECKS PASS", ck["A3"].value


def find_row(ws, prefix, col=1, start=1, stop=2000):
    for r in range(start, stop):
        v = ws.cell(row=r, column=col).value
        if isinstance(v, str) and v.startswith(prefix):
            return r
    raise KeyError(prefix)


DQ = []
for r in range(5, 25):
    row = [dq_ws.cell(row=r, column=c).value for c in range(1, 12)]
    assert row[0] == f"DQ{r - 4:02d}" and row[6] == "✔ OK", row[:7]
    DQ.append(dict(id=row[0], dim=row[1], cols=row[2], found=row[3], rows=row[4], done=row[7], why=row[8],
                   effect=row[9], qs=row[10]))

s1 = find_row(se, "Part 1") + 2
SENS = OrderedDict()
for key, r in zip(("base", "any", "invund", "zn", "dup", "noshow"), range(s1, s1 + 6)):
    SENS[key] = dict(label=se.cell(row=r, column=1).value, n=se.cell(row=r, column=2).value,
                     c=se.cell(row=r, column=3).value, rate=se.cell(row=r, column=4).value,
                     chg=se.cell(row=r, column=5).value)
s2 = find_row(se, "Part 2") + 2
LARGEST = {}
r = s2
while se.cell(row=r, column=1).value not in (None, ""):
    LARGEST[se.cell(row=r, column=1).value] = dict(change=se.cell(row=r, column=3).value,
                                                   where=se.cell(row=r, column=4).value)
    r += 1
s4 = find_row(se, "Part 4") + 2
SUPP = {}
r = s4
while se.cell(row=r, column=1).value not in (None, ""):
    SUPP[se.cell(row=r, column=1).value] = se.cell(row=r, column=2).value
    r += 1


def supp(prefix):
    hits = [v for k, v in SUPP.items() if k.startswith(prefix)]
    assert len(hits) == 1, prefix
    return hits[0]


BLOCKS = OrderedDict()
r = 5
while r < ov.max_row:
    v = ov.cell(row=r, column=1).value
    if isinstance(v, str) and v.startswith("Block "):
        key_title = v.split(": ", 1)[1]
        cats = []
        rr = r + 2
        while ov.cell(row=rr, column=1).value != "Total":
            cats.append(dict(label=str(ov.cell(row=rr, column=1).value), n=ov.cell(row=rr, column=2).value,
                             c=ov.cell(row=rr, column=3).value, rate=ov.cell(row=rr, column=4).value))
            rr += 1
        tot = dict(n=ov.cell(row=rr, column=2).value, c=ov.cell(row=rr, column=3).value)
        key = key_title.split(" ")[0].rstrip(":")
        BLOCKS[key] = dict(title=key_title, cats=cats, total=tot)
        r = rr + 1
    else:
        r += 1
assert len(BLOCKS) == 21, list(BLOCKS)
CAT = {k: {c["label"]: c for c in b["cats"]} for k, b in BLOCKS.items()}


def n(x):
    return f"{int(round(x)):,}"


def p1(x):
    return f"{x * 100:.1f}%"


def p2(x):
    return f"{x * 100:.2f}%"


def dmy(d):
    return f"{d.day} {d:%B %Y}"


def pts(x):
    return f"{x:+.2f}"


# ================================================================ document set-up
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.orientation = WD_ORIENT.PORTRAIT
for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, side, Cm(2.54))
TEXT_WIDTH = 21.0 - 2 * 2.54


def set_font(style_or_run, size=None, bold=None, italic=None, color=None):
    f = style_or_run.font
    f.name = FONT
    rpr = style_or_run.element.get_or_add_rPr() if hasattr(style_or_run.element, "get_or_add_rPr") else None
    if rpr is not None:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rfonts.set(qn(a), FONT)
    if size:
        f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    if color:
        f.color.rgb = RGBColor.from_string(color)


normal = doc.styles["Normal"]
set_font(normal, 12)
pf = normal.paragraph_format
pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
pf.space_after = Pt(6)
for lvl, size in ((1, 14), (2, 12)):
    st = doc.styles[f"Heading {lvl}"]
    set_font(st, size, bold=True, italic=False, color="000000")
    st.paragraph_format.space_before = Pt(12)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.keep_with_next = True
    st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
for name in ("List Bullet",):
    st = doc.styles[name]
    set_font(st, 12)
    st.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    st.paragraph_format.space_after = Pt(3)


def para(text, style=None, align=None, size=None, bold=None, italic=None, keep=False, space_after=None):
    p = doc.add_paragraph(style=style)
    run = p.add_run(text)
    set_font(run, size, bold, italic)
    if align is not None:
        p.paragraph_format.alignment = align
    if keep:
        p.paragraph_format.keep_with_next = True
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    return p


def rich_para(chunks, style=None):
    """chunks: list of (text, bold) pairs."""
    p = doc.add_paragraph(style=style)
    for text, bold in chunks:
        run = p.add_run(text)
        set_font(run, None, bold)
    return p


def heading(text, level=1):
    h = doc.add_heading(level=level)
    run = h.add_run(text)
    set_font(run, 14 if level == 1 else 12, bold=True, italic=False, color="000000")
    return h


def shade(cell, hex_fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcpr.append(shd)


def table(headers, rows, widths_cm, caption, source=None, align_right=()):
    assert abs(sum(widths_cm) - TEXT_WIDTH) < 0.05, sum(widths_cm)
    para(caption, align=WD_ALIGN_PARAGRAPH.LEFT, size=12, bold=True, keep=True, space_after=4)
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblpr = t._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr.append(layout)
    mar = OxmlElement("w:tblCellMar")  # 0.1 cm left/right cell padding, so narrow columns keep whole words
    for side in ("left", "right"):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), "57")
        m.set(qn("w:type"), "dxa")
        mar.append(m)
    tblpr.append(mar)
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    tblw.set(qn("w:w"), str(int(TEXT_WIDTH * 567)))
    tblw.set(qn("w:type"), "dxa")
    grid = t._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths_cm):
        gc.set(qn("w:w"), str(int(w * 567)))
    for i, row in enumerate([headers] + rows):
        tr = t.rows[i]
        trpr = tr._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit")  # keep every row whole on one page
        cs.set(qn("w:val"), "true")
        trpr.append(cs)
        if i == 0:  # header row repeats on every page the table reaches
            th = OxmlElement("w:tblHeader")
            th.set(qn("w:val"), "true")
            trpr.append(th)
        for j, (val, w) in enumerate(zip(row, widths_cm)):
            cell = tr.cells[j]
            cell.width = Cm(w)
            cell.text = ""
            p = cell.paragraphs[0]
            pfmt = p.paragraph_format
            pfmt.line_spacing_rule = WD_LINE_SPACING.SINGLE
            pfmt.space_after = Pt(0)
            pfmt.space_before = Pt(0)
            pfmt.alignment = (WD_ALIGN_PARAGRAPH.CENTER if i == 0 else
                              WD_ALIGN_PARAGRAPH.RIGHT if j in align_right else WD_ALIGN_PARAGRAPH.LEFT)
            lines = str(val).split("\n")
            for k, line in enumerate(lines):
                run = p.add_run(line)
                set_font(run, 9, bold=(i == 0 or (k == 0 and len(lines) > 1 and j == 0)),
                         color="FFFFFF" if i == 0 else None)
                if k < len(lines) - 1:
                    run.add_break()
            if i == 0:
                shade(cell, NAVY)
            elif i % 2 == 0:
                shade(cell, "F2F2F2")
    if source:
        para(source, align=WD_ALIGN_PARAGRAPH.LEFT, size=10, italic=True, space_after=10)
    return t


# ================================================================ 1. title
para("Data quality notes: BUSM1042 Assignment 3 hotel bookings dataset", align=WD_ALIGN_PARAGRAPH.CENTER, size=16,
     bold=True, space_after=2)
para("Prepared for the group report; every count is linked to BUSM1042_A3_Group_Clean_Data.xlsx",
     align=WD_ALIGN_PARAGRAPH.CENTER, size=12, italic=True, space_after=12)

# ================================================================ 2. how the data was checked
heading("1. How the data was checked")
para("The team worked on a copy of the supplied file and kept the original unchanged. Each column was checked "
     "against the variable definitions in the Data in Brief paper that published the dataset (Antonio, de Almeida & "
     "Nunes 2019): what it means, what kind of values it holds and anything that looks odd. Quality was judged by "
     "fitness for use, so a problem matters only if it could change an answer to one of our five business questions "
     "(Vidgen, Kirshner & Tan 2019, p. 58). Suspect values were flagged rather than deleted or guessed, because "
     "deleting cases can bias results (Vidgen, Kirshner & Tan 2019, p. 63). As the scenario requires, every booking "
     "is treated as an Australian tourist's, so the country column is never used to cut the data.")

# ================================================================ 3. the issue table
heading("2. Data quality issues and treatment")
para("The table lists the 20 issues found, in the order of the workbook's Data_Quality_Log sheet. No row was "
     "deleted: problems are handled with flags, recoded copies of columns and sensitivity checks.")
rows = []
for d in DQ:
    rows_txt = n(d["rows"]) if isinstance(d["rows"], (int, float)) else str(d["rows"])
    rows.append([f"{d['id']}\n{d['dim']}", d["found"], rows_txt, f"{d['done']}. {d['why']}.", d["effect"], d["qs"]])
table(["Quality dimension", "What was found", "Rows affected", "What was done and why", "Effect on results",
       "Questions affected"], rows, [2.3, 4.2, 1.5, 3.3, 2.75, 1.87],
      "Table X: Data quality checks and treatment",
      f"Source: Group analysis of the supplied hotel bookings dataset (n = {n(SENS['base']['n'])}); quality dimensions "
      "from Vidgen, Kirshner and Tan (2019, pp. 57–58).", align_right=(2,))

# ================================================================ 4. sensitivity
heading("3. How much could these choices change our results?")
b, a = SENS["base"], SENS["any"]
grp = LARGEST["customer_type"]
para(f"Excluding all {n(b['n'] - a['n'])} flagged rows (invalid, undefined or outlier values) moves the "
     f"overall cancellation rate from {p2(b['rate'])} to {p2(a['rate'])}, a change of only {pts(a['chg'])} points, so "
     f"these flags do not change any overall conclusion. Two choices do matter. Removing the "
     f"{n(b['n'] - SENS['dup']['n'])} duplicate rows would cut the rate from {p1(b['rate'])} to "
     f"{p1(SENS['dup']['rate'])}, and counting no-shows as not cancelled would lower it to "
     f"{p1(SENS['noshow']['rate'])}. Every question should therefore state that cancellations include no-shows and "
     f"check its main pattern without duplicates (dup_flag = 0). Inside single tables, a few small categories move "
     f"more, for example customer type {grp['where']}, which moves by {grp['change']:.2f} points, and teammates using "
     f"those variables should say so.")
srows = []
labels = {"base": "All bookings (baseline)", "any": "Excluding all flagged rows (flag_any_quality)",
          "invund": "Excluding invalid and undefined rows only", "zn": "Excluding zero-night stays",
          "dup": "Excluding duplicates (dup_flag = 1)", "noshow": "Counting only 'Canceled' (no-shows not counted)"}
for k, v in SENS.items():
    srows.append([labels[k], n(v["n"]), p2(v["rate"]), "baseline" if k == "base" else pts(v["chg"])])
table(["Test", "Bookings used", "Overall cancellation rate", "Change (points)"], srows, [8.42, 2.5, 2.5, 2.5],
      "Table Y: How much each decision changes the overall cancellation rate",
      "Source: Sensitivity sheet of BUSM1042_A3_Group_Clean_Data.xlsx. Points = percentage points.",
      align_right=(1, 2, 3))

# ================================================================ 5. notes for each business question
heading("4. Notes for each business question")
hotel, dep, seg, chn = CAT["hotel_name"], CAT["deposit_type"], CAT["market_segment"], CAT["distribution_channel"]
lead, rep = CAT["lead_band"], CAT["is_repeated_guest"]
invalid_dates = CAT["arrival_month"]["Invalid date"]["n"]
QNOTES = [
    ("Q1 (size of the problem; City vs Resort; over time). ",
     f"Use date_valid = 1 for any monthly analysis, because {n(invalid_dates)} bookings have an impossible date. "
     f"Arrivals run from {dmy(supp('Earliest valid arrival date'))} to {dmy(supp('Latest valid arrival date'))}, so "
     f"2022 and 2024 are part-years: compare rates or months, not yearly totals. Never split or filter by country. City Hotel bookings "
     f"were cancelled {p1(hotel['City Hotel']['rate'])} of the time, against {p1(hotel['Resort Hotel']['rate'])} for "
     f"Resort Hotel."),
    ("Q2 (lead time). ",
     f"Use the six lead-time bands. The Over 365 band is small ({p1(supp('Over 365 band: share of all bookings'))} "
     f"of bookings) and {p1(supp('Over 365 band: share that is Non Refund'))} of it is non-refundable, so its high "
     f"rate ({p1(lead['Over 365']['rate'])}) partly reflects deposit policy rather than time alone."),
    ("Q3 (no-deposit bookings; channels and segments). ",
     f"The {n(DQ[7]['rows'])} bookings with an Undefined segment or channel are flagged. Aviation "
     f"({n(seg['Aviation']['n'])} bookings) and Complementary ({n(seg['Complementary']['n'])}) are small segments, "
     f"and GDS has only {n(chn['GDS']['n'])} bookings, so report their rates with care."),
    ("Q4 (deposit type × segment). ",
     f"Non Refund bookings were cancelled {p1(dep['Non Refund']['rate'])} of the time, against "
     f"{p1(dep['No Deposit']['rate'])} for No Deposit. Refundable has only {n(dep['Refundable']['n'])} bookings, so "
     f"its rates ({p1(dep['Refundable']['rate'])} overall) are unreliable once it is split by segment; say so."),
    ("Q5 (booking history). ",
     f"Values of previous_cancellations from 11 to 26 ({n(supp('Bookings with 11 to 26'))} rows) are probably "
     f"repeated group records: {p1(supp('…as a share of those bookings'))} of them are Groups or Offline TA/TO "
     f"bookings. Treat them as one '2+' band (prev_canc_band). Repeat guests cancelled {p1(rep['Yes']['rate'])} of "
     f"their bookings, against {p1(rep['No']['rate'])} for other guests."),
    ("For everyone. ",
     f"Parking and room changes are 'use with caution' (DQ20). {n(supp('…of which cancelled'))} of "
     f"{n(supp('Bookings with 1+ parking spaces'))} bookings with a parking space were cancelled, and only "
     f"{p1(supp('…their cancellation rate'))} of bookings with a changed room type, because these details are often "
     f"set at check-in. Do not present them as reasons guests cancel."),
]
for lead_txt, body in QNOTES:
    rich_para([(lead_txt, True), (body, False)])

# ================================================================ 6. what stands out
heading("5. What stands out (to help write the business questions)")
para("Each bullet compares the highest and lowest cancellation rate within one variable, using only categories with "
     "30 or more bookings; bullets run from the largest gap to the smallest. Rates come from the OneVar_Summary sheet.")
NICE = {"hotel_name": "Hotel", "lead_band": "Lead time", "deposit_type": "Deposit type",
        "market_segment": "Market segment", "distribution_channel": "Distribution channel",
        "customer_type": "Customer type", "is_repeated_guest": "Repeat guest", "prev_canc_band": "Previous cancellations",
        "prev_notcanc_band": "Previous bookings not cancelled", "meal_clean": "Meal (Undefined counted as SC)",
        "changes_band": "Booking changes", "requests_band": "Special requests", "waiting_band": "Days on the waiting list",
        "nights_band": "Total nights", "adr_band": "Average daily rate", "with_children": "Children or babies",
        "arrival_month": "Arrival month (valid dates only)", "arrival_year": "Arrival year",
        "parking_band": "Parking spaces", "room_changed": "Room type changed", "reserved_room_type": "Reserved room type"}
WORDS = {  # plain-English category names for the bullets
    "lead_band": {lab: f"{lab} days" for lab in ("0–7", "8–30", "31–90", "91–180", "181–365", "Over 365")},
    "prev_canc_band": {"0": "guests with no previous cancellation", "1": "guests with one previous cancellation",
                       "2+": "guests with two or more previous cancellations"},
    "prev_notcanc_band": {"0": "guests with no earlier booking that was kept",
                          "1+": "guests with one or more earlier bookings that were kept"},
    "changes_band": {"0": "no booking changes", "1": "one change", "2": "two changes", "3+": "three or more changes"},
    "requests_band": {"0": "no special requests", "1": "one request", "2": "two requests",
                      "3+": "three or more requests"},
    "waiting_band": {"0": "no days on the waiting list", "1+": "one or more days waiting"},
    "nights_band": {"0": "0 nights", "1": "1 night", "2–3": "2–3 nights", "4–7": "4–7 nights", "8+": "8 or more nights"},
    "adr_band": {"0 or less": "a rate of 0 or less", "0.01–50": "a rate of 0.01–50", "50.01–100": "a rate of 50.01–100",
                 "100.01–150": "a rate of 100.01–150", "150.01–200": "a rate of 150.01–200",
                 "Over 200": "a rate over 200"},
    "parking_band": {"0": "no parking space", "1+": "one or more parking spaces"},
    "is_repeated_guest": {"Yes": "repeat guests", "No": "other guests"},
    "with_children": {"Yes": "with children or babies", "No": "adults only"},
    "room_changed": {"Same": "same room type", "Different": "changed room type"},
    "arrival_month": dict(zip(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
                              ["January", "February", "March", "April", "May", "June", "July", "August", "September",
                               "October", "November", "December"])),
    "reserved_room_type": {c: f"type {c}" for c in "ABCDEFGHIJKLMNOP"},
}
EXTRA = {"deposit_type": "; check whether this is real behaviour or a recording practice (report Section 6.4)",
         "arrival_year": "; 2022 and 2024 are part-years, so compare rates only",
         "parking_band": "; use with caution, as this may be recorded at check-in (DQ20)",
         "room_changed": "; use with caution, as this may be recorded at check-in (DQ20)",
         "reserved_room_type": "; use with caution, as this may be recorded at check-in (DQ20)"}


def show(key, lab):
    return WORDS.get(key, {}).get(lab, lab)


bullets = []
for key, blk in BLOCKS.items():
    ok = [c for c in blk["cats"] if c["n"] >= 30 and c["label"] != "Invalid date"]
    hi = max(ok, key=lambda c: c["rate"])
    lo = min(ok, key=lambda c: c["rate"])
    small = [c for c in blk["cats"] if c["n"] < 30]
    txt = (f"{NICE[key]}: {show(key, hi['label'])} {p1(hi['rate'])} vs {show(key, lo['label'])} {p1(lo['rate'])} "
           f"(a gap of {(hi['rate'] - lo['rate']) * 100:.1f} points)")
    if key == "deposit_type":
        txt += f"; No Deposit, the largest group, {p1(CAT[key]['No Deposit']['rate'])}"
    txt += EXTRA.get(key, "")
    if small:
        txt += "; unreliable (n < 30): " + ", ".join(f"{c['label']} ({n(c['n'])} bookings, {p1(c['rate'])})"
                                                     for c in small)
    bullets.append(((hi["rate"] - lo["rate"]), txt + "."))
for _, txt in sorted(bullets, key=lambda x: -x[0]):
    para(txt, style="List Bullet")

# ================================================================ 7. limitations
heading("6. Limitations")
for txt in ("The file has no cancellation date (the paper's ReservationStatusDate is not included), so the notice "
            "hotels receive before a cancellation cannot be measured.",
            "The patterns are associations, not causes: a variable linked with a higher cancellation rate does not "
            "on its own make guests cancel.",
            "Some fields may be recorded at check-in, such as parking spaces and room changes (DQ20), so they partly "
            "describe the outcome rather than explain it."):
    para(txt, style="List Bullet")

# ================================================================ 8. references
heading("7. Reference list")
p = doc.add_paragraph()
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.left_indent = Cm(1.0)
p.paragraph_format.first_line_indent = Cm(-1.0)
for text, it in (("Antonio, N, de Almeida, A & Nunes, L 2019, 'Hotel booking demand datasets', ", False),
                 ("Data in Brief", True),
                 (f", vol. 22, pp. 41–49, viewed {VIEWED}, <https://doi.org/10.1016/j.dib.2018.11.126>.", False)):
    set_font(p.add_run(text), None, None, it)
p = doc.add_paragraph()
p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.left_indent = Cm(1.0)
p.paragraph_format.first_line_indent = Cm(-1.0)
for text, it in (("Vidgen, R, Kirshner, S & Tan, F 2019, ", False), ("Business analytics: a management approach", True),
                 (", Red Globe Press, London.", False)):
    set_font(p.add_run(text), None, None, it)
para("Note for the group: if the report also cites other Antonio, de Almeida & Nunes 2019 papers, label them 2019a, "
     "2019b and 2019c consistently. In Mark's sections the Data in Brief paper is 2019c.",
     align=WD_ALIGN_PARAGRAPH.LEFT, size=10, italic=True)

# page numbers in the footer
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = fp.add_run()
for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
    if tag:
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), tag)
        run._r.append(fc)
    else:
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = text
        run._r.append(it)
set_font(run, 10)

full_text = "\n".join(p.text for p in doc.paragraphs) + "\n".join(
    c.text for t in doc.tables for row in t.rows for c in row.cells)
assert "—" not in full_text, "em dash found"
doc.save(OUT)
print(f"Saved {OUT}")
