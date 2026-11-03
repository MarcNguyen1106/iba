"""Build BUSM1042_A3_Group_Clean_Data.xlsx: the shared, audited dataset for business questions Q1-Q5.

Usage:  python3 -I build_group_workbook.py <IBA-A3-Hotel_bookings_dataset-T32026.xlsx> <output.xlsx>

The source file is only read. No row is deleted and no source value is overwritten: Raw_Data is an exact copy,
Clean_Data converts only the text codes "NULL"/"NA" to empty cells and adds derived columns as row formulas.
Every count, rate and share in a summary sheet is an Excel formula; typed numbers appear only in columns
labelled "Expected", which the checks compare against. The one Python-computed column is Clean_Data!dup_flag.
"""
import re
import sys
import time
from collections import Counter
from datetime import datetime

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableStyleInfo

SRC, OUT = sys.argv[1], sys.argv[2]
BUILD_DATE = "3 November 2026"
N = 119390
LAST = N + 1  # last data row on Raw_Data / Clean_Data

# ---------------------------------------------------------------- style
FONT = "Calibri"
NAVY, GRID, BAND, TEAL, LGREY, AMBER = "1F3A5F", "BFBFBF", "F2F2F2", "CCECE8", "D9D9D9", "FCE4D6"
F_BODY = Font(name=FONT, size=11)
F_BOLD = Font(name=FONT, size=11, bold=True)
F_TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
F_TTITLE = Font(name=FONT, size=12, bold=True, color=NAVY)
F_HDR = Font(name=FONT, size=11, bold=True, color="FFFFFF")
F_NOTE = Font(name=FONT, size=10, italic=True, color="595959")
F_SMALL = Font(name=FONT, size=10, color="404040")
F_HELP = Font(name=FONT, size=9, italic=True, color="808080")
F_EXP = Font(name=FONT, size=11, color="0000FF")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_BAND = PatternFill("solid", fgColor=BAND)
FILL_GREY = PatternFill("solid", fgColor=LGREY)
FILL_TEAL = PatternFill("solid", fgColor=TEAL)
FILL_AMBER = PatternFill("solid", fgColor=AMBER)
THIN = Side(style="thin", color=GRID)
B_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
B_TOTAL = Border(left=THIN, right=THIN, top=Side(style="medium", color="404040"), bottom=THIN)
A_HDR = Alignment(horizontal="center", vertical="center", wrap_text=True)
A_WRAP = Alignment(vertical="top", wrap_text=True)
A_LEFT = Alignment(vertical="center")
A_NUM = Alignment(horizontal="right", vertical="center")
A_CTR = Alignment(horizontal="center", vertical="center")
FMT_N, FMT_P, FMT_P2, FMT_PT = "#,##0", "0.0%", "0.00%", "+0.00;-0.00;0.00"
OK, BAD = "✔ OK", "✘ CHECK"
GREEN_FILL, GREEN_FONT = PatternFill("solid", fgColor="C6EFCE"), Font(name=FONT, color="006100", bold=True)
RED_FILL, RED_FONT = PatternFill("solid", fgColor="FFC7CE"), Font(name=FONT, color="9C0006", bold=True)
CAUTION = "use with caution: may be recorded at check-in"


def lit(text):
    """Excel string literal; long text is split because one literal may hold at most 255 characters."""
    chunks = [text[i:i + 200] for i in range(0, len(text), 200)] or [""]
    return "&".join('"' + c.replace('"', '""') + '"' for c in chunks)


def F(*parts):
    """Formula text from parts: a str is literal text; (expr, fmt) is TEXT(expr, fmt); ("=", expr) is raw."""
    out = []
    for p in parts:
        if isinstance(p, str):
            if p:
                out.append(lit(p))
        elif p[0] == "=":
            out.append(f"({p[1]})")
        else:
            out.append(f'TEXT({p[0]},"{p[1]}")')
    return "=" + "&".join(out)


def put(ws, ref, value, font=F_BODY, fmt=None, align=None, fill=None, border=None):
    c = ws[ref]
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


def header_row(ws, row, left, headers, height=None):
    for j, h in enumerate(headers):
        c = ws.cell(row=row, column=left + j, value=h)
        c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, A_HDR, B_ALL
    if height:
        ws.row_dimensions[row].height = height


def body_cell(ws, row, col, value, fmt=None, band=False, bold=False, total=False, wrap=False, align=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = F_BOLD if (bold or total) else F_BODY
    if fmt and not (fmt == "@" and isinstance(value, str) and value.startswith("=")):
        c.number_format = fmt
    c.alignment = align or (A_WRAP if wrap else (A_NUM if fmt not in (None, "@") else A_LEFT))
    c.border = B_TOTAL if total else B_ALL
    if band and not total:
        c.fill = FILL_BAND
    return c


def title(ws, ref, text):
    put(ws, ref, text, font=F_TITLE)
    ws.row_dimensions[ws[ref].row].height = 22


def ttitle(ws, ref, text):
    put(ws, ref, text, font=F_TTITLE)
    ws.row_dimensions[ws[ref].row].height = 18


def presentation(ws, tab, landscape=True, area=None):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins.left = ws.page_margins.right = 0.5
    if area:
        ws.print_area = area


def match_rules(ws, rng):
    first = rng.split(":")[0].replace("$", "")
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{OK}"'], fill=GREEN_FILL, font=GREEN_FONT))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="{BAD}"'], fill=RED_FILL, font=RED_FONT))


def bool_rules(ws, rng):
    first = rng.split(":")[0].replace("$", "")
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"{first}=TRUE"], fill=GREEN_FILL, font=GREEN_FONT))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"{first}<>TRUE"], fill=RED_FILL, font=RED_FONT))


t0 = time.time()
# ================================================================ 1. load the source exactly as stored
print("Reading the source exactly as stored ...", flush=True)
src_wb = load_workbook(SRC, read_only=True)
src_rows = [list(r) for r in src_wb["hotels"].iter_rows(values_only=True)]
src_wb.close()
cols = src_rows[0]
raw_rows = src_rows[1:]
assert len(cols) == 29 and len(raw_rows) == N, (len(cols), len(raw_rows))
COL = {name: get_column_letter(i + 1) for i, name in enumerate(cols)}
EXPECTED_ORDER = ["booking_order_id", "hotel_name", "is_canceled", "lead_time", "arrival_date",
                  "stays_in_weekend_nights", "stays_in_week_nights", "adults", "children", "babies", "meal",
                  "country", "market_segment", "distribution_channel", "is_repeated_guest",
                  "previous_cancellations", "previous_bookings_not_canceled", "reserved_room_type",
                  "assigned_room_type", "booking_changes", "deposit_type", "agent", "company",
                  "days_in_waiting_list", "customer_type", "Average_Daily_Rate", "required_car_parking_spaces",
                  "total_of_special_requests", "reservation_status"]
assert cols == EXPECTED_ORDER, cols  # the formula letters below assume A booking_order_id ... AC reservation_status
assert COL["is_canceled"] == "C" and COL["Average_Daily_Rate"] == "Z" and COL["reservation_status"] == "AC"

# pandas copy with every value kept as stored (the texts "NULL" and "NA" are NOT turned into blanks)
df = pd.read_excel(SRC, sheet_name="hotels", engine="openpyxl", keep_default_na=False, dtype=object)
assert df.shape == (N, 29) and list(df.columns) == cols
assert df["booking_order_id"].tolist() == [r[0] for r in raw_rows]
dup = df.duplicated(subset=cols[1:], keep="first").astype(int).tolist()  # DQ17

# ================================================================ Step 1: profile every column (before cleaning)
NUMERIC = ["lead_time", "stays_in_weekend_nights", "stays_in_week_nights", "adults", "children", "babies",
           "previous_cancellations", "previous_bookings_not_canceled", "booking_changes", "days_in_waiting_list",
           "Average_Daily_Rate", "required_car_parking_spaces", "total_of_special_requests"]
MEANING = {  # Antonio, de Almeida & Nunes (2019), Table 1, pp. 43-45, unless stated
    "booking_order_id": "Booking identifier added to this file (not in the paper); one per booking.",
    "hotel_name": "Hotel type. The paper publishes two datasets: H1, a resort hotel (Resort Hotel), and H2, a city "
                  "hotel (City Hotel).",
    "is_canceled": "IsCanceled: whether the booking was cancelled (1) or not (0). Coded Yes/No in this file.",
    "lead_time": "LeadTime: days between the date the booking was entered into the hotel's system (PMS) and the "
                 "arrival date.",
    "arrival_date": "Arrival date. The paper has separate year, month, week number and day of month for arrivals "
                    "from 1 July 2015 to 31 August 2017; this file combines them into one date, seven years later.",
    "stays_in_weekend_nights": "StaysInWeekendNights: weekend nights (Saturday or Sunday) the guest stayed or "
                               "booked to stay.",
    "stays_in_week_nights": "StaysInWeekNights: week nights (Monday to Friday) the guest stayed or booked to stay.",
    "adults": "Adults: number of adults.",
    "children": "Children: number of children (paying and non-paying).",
    "babies": "Babies: number of babies.",
    "meal": "Meal: meal package booked. Undefined/SC = no meal package; BB = bed and breakfast; HB = half board "
            "(breakfast and usually dinner); FB = full board (breakfast, lunch and dinner).",
    "country": "Country: guest's country of origin in ISO country codes. Origin, not citizenship; never used to "
               "filter (scenario rule).",
    "market_segment": "MarketSegment: market segment designation (TA = travel agents, TO = tour operators).",
    "distribution_channel": "DistributionChannel: booking distribution channel (TA = travel agents, TO = tour "
                            "operators).",
    "is_repeated_guest": "IsRepeatedGuest: whether the booking name was a repeated guest (1) or not (0). Coded "
                         "Yes/No in this file.",
    "previous_cancellations": "PreviousCancellations: the customer's earlier bookings that were cancelled before "
                              "this booking.",
    "previous_bookings_not_canceled": "PreviousBookingsNotCanceled: the customer's earlier bookings that were not "
                                      "cancelled.",
    "reserved_room_type": "ReservedRoomType: code of the room type reserved (a code instead of a name, for "
                          "anonymity).",
    "assigned_room_type": "AssignedRoomType: code of the room type assigned; it can differ from the reserved "
                          "type, for example after overbooking or at the customer's request.",
    "booking_changes": "BookingChanges: changes made to the booking between entry in the PMS and check-in or "
                       "cancellation.",
    "deposit_type": "DepositType: No Deposit = none; Non Refund = deposit of the total stay cost; Refundable = "
                    "deposit below the total stay cost.",
    "agent": "Agent: ID of the travel agency that made the booking. \"NULL\" means not applicable (no agent), not "
             "missing (paper, p. 45).",
    "company": "Company: ID of the company or entity that made or pays for the booking. \"NULL\" means not "
               "applicable (paper, p. 45).",
    "days_in_waiting_list": "DaysInWaitingList: days the booking was on the waiting list before it was confirmed.",
    "customer_type": "CustomerType: Contract, Group, Transient (not part of a group or contract) or "
                     "Transient-Party (transient but linked to at least one other transient booking).",
    "Average_Daily_Rate": "ADR: average daily rate, the sum of all lodging transactions divided by the number of "
                          "staying nights.",
    "required_car_parking_spaces": "RequiredCarParkingSpaces: car parking spaces the customer required.",
    "total_of_special_requests": "TotalOfSpecialRequests: special requests made (for example twin bed or high "
                                 "floor).",
    "reservation_status": "ReservationStatus: last status: Canceled (cancelled by the customer), Check-Out "
                          "(checked in and departed) or No-Show (did not check in).",
}
assert set(MEANING) == set(cols)


def check_profile(frame):
    """Assert the expected profile facts from the brief (A1). Any failure stops the build."""
    col = {c: frame[c] for c in cols}
    vc = {c: Counter(col[c]) for c in cols}
    assert frame.shape == (N, 29) and col["booking_order_id"].nunique() == N
    assert vc["hotel_name"] == {"City Hotel": 79330, "Resort Hotel": 40060}
    assert vc["is_canceled"] == {"Yes": 44224, "No": 75166}
    assert vc["reservation_status"] == {"Canceled": 43017, "No-Show": 1207, "Check-Out": 75166}
    lt = pd.to_numeric(col["lead_time"])
    assert (lt.min(), lt.median(), round(lt.mean(), 2), lt.max(), int((lt == 0).sum())) == (0, 69, 104.01, 737, 6345)
    dates = [v for v in col["arrival_date"] if isinstance(v, datetime)]
    texts = Counter(v for v in col["arrival_date"] if not isinstance(v, datetime))
    assert len(dates) == 119294 and min(dates) == datetime(2022, 7, 1) and max(dates) == datetime(2024, 8, 31)
    assert texts == {"29 February 2023": 96}
    assert vc["children"] == {0: 110796, 1: 4861, 2: 3652, 3: 76, 10: 1, "NA": 4}
    ad = pd.to_numeric(col["adults"])
    assert int((ad == 0).sum()) == 403 and ad.max() == 55
    assert vc["babies"] == {0: 118473, 1: 900, 2: 15, 9: 1, 10: 1}
    assert vc["meal"] == {"BB": 92310, "HB": 14463, "SC": 10650, "Undefined": 1169, "FB": 798}
    assert len(vc["country"]) == 178 and vc["country"]["NULL"] == 488
    assert vc["country"]["CN"] == 1279 and vc["country"]["CHN"] == 999
    assert vc["market_segment"] == {"Online TA": 56477, "Offline TA/TO": 24219, "Groups": 19811, "Direct": 12606,
                                    "Corporate": 5295, "Complementary": 743, "Aviation": 237, "Undefined": 2}
    assert vc["distribution_channel"] == {"TA/TO": 97870, "Direct": 14645, "Corporate": 6677, "GDS": 193,
                                          "Undefined": 5}
    assert vc["is_repeated_guest"] == {"No": 115580, "Yes": 3810}
    assert vc["deposit_type"] == {"No Deposit": 104641, "Non Refund": 14587, "Refundable": 162}
    assert vc["agent"]["NULL"] == 16340 and vc["company"]["NULL"] == 112593
    assert vc["customer_type"] == {"Transient": 89613, "Transient-Party": 25124, "Contract": 4076, "Group": 577}
    adr = pd.to_numeric(col["Average_Daily_Rate"])
    assert (adr.min(), adr.median(), adr.max(), int((adr == 0).sum())) == (-6.38, 94.575, 5400, 1959)
    assert vc["required_car_parking_spaces"] == {0: 111974, 1: 7383, 2: 28, 3: 3, 8: 2}
    for c, lo, hi in (("total_of_special_requests", 0, 5), ("booking_changes", 0, 21),
                      ("days_in_waiting_list", 0, 391), ("previous_cancellations", 0, 26),
                      ("previous_bookings_not_canceled", 0, 72)):
        x = pd.to_numeric(col[c])
        assert (x.min(), x.max()) == (lo, hi), c


check_profile(df)
print("Profile facts: all expected values confirmed.", flush=True)

profile = []
for c in cols:
    s = df[c]
    types = Counter(type(v).__name__ for v in s)
    profile.append(dict(name=c, types=", ".join(f"{k} {v:,}" for k, v in types.most_common()),
                        distinct=int(s.nunique()), measure=c in NUMERIC,
                        top=[k for k, _ in Counter(s).most_common(5)]))

# ================================================================ cleaning step: NULL / NA -> empty cell
NULL_TEXT = {"children": "NA", "country": "NULL", "agent": "NULL", "company": "NULL"}
null_idx = {cols.index(k): v for k, v in NULL_TEXT.items()}
clean_rows = []
converted = 0
for r in raw_rows:
    r = list(r)
    for i, token in null_idx.items():
        if r[i] == token:
            r[i] = None
            converted += 1
    clean_rows.append(r)
assert converted == 488 + 16340 + 112593 + 4 == 129425, converted

wb = Workbook()
wb.calculation.fullCalcOnLoad = True
wb.calculation.calcMode = "auto"
ws_readme = wb.active
ws_readme.title = "README"
ws_raw = wb.create_sheet("Raw_Data")
ws_prof = wb.create_sheet("Column_Profile")
ws_cl = wb.create_sheet("Clean_Data")
ws_dq = wb.create_sheet("Data_Quality_Log")
ws_sens = wb.create_sheet("Sensitivity")
ws_ov = wb.create_sheet("OneVar_Summary")
ws_ch = wb.create_sheet("Checks")
DATE_COL = cols.index("arrival_date")


def write_data_sheet(ws, rows, extra_headers=()):
    ws.append(list(cols) + list(extra_headers))
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows(min_row=2, max_row=LAST, min_col=DATE_COL + 1, max_col=DATE_COL + 1):
        if isinstance(row[0].value, datetime):
            row[0].number_format = "dd/mm/yyyy"
    for j in range(1, 30):
        h = ws.cell(row=1, column=j)
        h.font, h.fill, h.alignment = F_BOLD, FILL_GREY, A_HDR
        ws.column_dimensions[get_column_letter(j)].width = max(11, min(30, len(str(h.value)) + 3))
    ws.column_dimensions["A"].width = 17
    ws.column_dimensions[COL["arrival_date"]].width = 17
    ws.freeze_panes = "A2"


# ================================================================ Raw_Data: the untouched original
print("Writing Raw_Data ...", flush=True)
write_data_sheet(ws_raw, raw_rows)
ws_raw.sheet_properties.tabColor = "2E75B6"
ws_raw.auto_filter.ref = f"A1:AC{LAST}"
ws_raw["A1"].comment = Comment("Untouched original: an exact copy of the source sheet 'hotels'. Do not edit.",
                               "BUSM1042 group")
ws_raw.protection.sheet = True          # no password
ws_raw.protection.autoFilter = False    # filtering stays allowed on the protected sheet
ws_raw.protection.sort = False
del raw_rows


def R(field):
    """Exact range of a source column on Raw_Data."""
    return f"Raw_Data!${COL[field]}$2:${COL[field]}${LAST}"


# ================================================================ Clean_Data (tblClean)
C = COL
DERIVED = [  # name, formula for row r (plain A1 references), kind ("d" derived, "f" flag)
    ("cancel_flag", lambda r: f'=IF(C{r}="Yes",1,0)', "d"),
    ("repeat_flag", lambda r: f'=IF(O{r}="Yes",1,0)', "d"),
    ("date_valid", lambda r: f"=IF(ISNUMBER(E{r}),1,0)", "d"),
    ("arrival_year", lambda r: f'=IF({X["date_valid"]}{r}=1,YEAR(E{r}),"Invalid date")', "d"),
    ("arrival_month", lambda r: f'=IF({X["date_valid"]}{r}=1,MONTH(E{r}),"Invalid date")', "d"),
    ("arrival_month_name", lambda r: f'=IF({X["date_valid"]}{r}=1,TEXT(E{r},"mmm"),"Invalid date")', "d"),
    ("lead_band", lambda r: (f'=IF(D{r}<=7,"0–7",IF(D{r}<=30,"8–30",IF(D{r}<=90,"31–90",IF(D{r}<=180,"91–180",'
                             f'IF(D{r}<=365,"181–365","Over 365")))))'), "d"),
    ("lead_band_order", lambda r: f"=IF(D{r}<=7,1,IF(D{r}<=30,2,IF(D{r}<=90,3,IF(D{r}<=180,4,IF(D{r}<=365,5,6)))))",
     "d"),
    ("total_nights", lambda r: f"=F{r}+G{r}", "d"),
    ("nights_band", lambda r: (f'=IF({X["total_nights"]}{r}=0,"0",IF({X["total_nights"]}{r}=1,"1",'
                               f'IF({X["total_nights"]}{r}<=3,"2–3",IF({X["total_nights"]}{r}<=7,"4–7","8+"))))'), "d"),
    ("nights_band_order", lambda r: (f'=IF({X["total_nights"]}{r}=0,1,IF({X["total_nights"]}{r}=1,2,'
                                     f'IF({X["total_nights"]}{r}<=3,3,IF({X["total_nights"]}{r}<=7,4,5))))'), "d"),
    ("guests_total", lambda r: f"=H{r}+N(I{r})+J{r}", "d"),
    ("with_children", lambda r: f'=IF(N(I{r})+J{r}>0,"Yes","No")', "d"),
    ("meal_clean", lambda r: f'=IF(K{r}="Undefined","SC",K{r})', "d"),
    ("country_clean", lambda r: f'=IF(L{r}="","Unknown",L{r})', "d"),
    ("agent_clean", lambda r: f'=IF(V{r}="","None",V{r})', "d"),
    ("company_clean", lambda r: f'=IF(W{r}="","None",W{r})', "d"),
    ("prev_canc_band", lambda r: f'=IF(P{r}=0,"0",IF(P{r}=1,"1","2+"))', "d"),
    ("prev_canc_band_order", lambda r: f"=IF(P{r}=0,1,IF(P{r}=1,2,3))", "d"),
    ("prev_notcanc_band", lambda r: f'=IF(Q{r}=0,"0","1+")', "d"),
    ("prev_notcanc_band_order", lambda r: f"=IF(Q{r}=0,1,2)", "d"),
    ("changes_band", lambda r: f'=IF(T{r}=0,"0",IF(T{r}=1,"1",IF(T{r}=2,"2","3+")))', "d"),
    ("changes_band_order", lambda r: f"=IF(T{r}=0,1,IF(T{r}=1,2,IF(T{r}=2,3,4)))", "d"),
    ("requests_band", lambda r: f'=IF(AB{r}=0,"0",IF(AB{r}=1,"1",IF(AB{r}=2,"2","3+")))', "d"),
    ("requests_band_order", lambda r: f"=IF(AB{r}=0,1,IF(AB{r}=1,2,IF(AB{r}=2,3,4)))", "d"),
    ("parking_band", lambda r: f'=IF(AA{r}=0,"0","1+")', "d"),
    ("parking_band_order", lambda r: f"=IF(AA{r}=0,1,2)", "d"),
    ("waiting_band", lambda r: f'=IF(X{r}=0,"0","1+")', "d"),
    ("waiting_band_order", lambda r: f"=IF(X{r}=0,1,2)", "d"),
    ("adr_band", lambda r: (f'=IF(Z{r}<=0,"0 or less",IF(Z{r}<=50,"0.01–50",IF(Z{r}<=100,"50.01–100",'
                            f'IF(Z{r}<=150,"100.01–150",IF(Z{r}<=200,"150.01–200","Over 200")))))'), "d"),
    ("adr_band_order", lambda r: f"=IF(Z{r}<=0,1,IF(Z{r}<=50,2,IF(Z{r}<=100,3,IF(Z{r}<=150,4,IF(Z{r}<=200,5,6)))))",
     "d"),
    ("room_changed", lambda r: f'=IF(R{r}=S{r},"Same","Different")', "d"),
    ("flag_invalid", lambda r: (f'=IF(OR(I{r}="",{X["guests_total"]}{r}=0,Z{r}<0,Z{r}=ADR_MAX),1,0)'), "f"),
    ("flag_undefined", lambda r: f'=IF(OR(M{r}="Undefined",N{r}="Undefined"),1,0)', "f"),
    ("flag_outlier", lambda r: (f'=IF(OR(AND(H{r}=0,{X["guests_total"]}{r}>0),H{r}>=10,N(I{r})>=10,J{r}>=9,'
                                f'AA{r}>=8),1,0)'), "f"),
    ("flag_zero_rate", lambda r: f"=IF(Z{r}=0,1,0)", "f"),
    ("flag_zero_nights", lambda r: f'=IF({X["total_nights"]}{r}=0,1,0)', "f"),
    ("flag_any_quality", lambda r: (f'=IF(OR({X["flag_invalid"]}{r}=1,{X["flag_undefined"]}{r}=1,'
                                    f'{X["flag_outlier"]}{r}=1),1,0)'), "f"),
]
DNAMES = [d[0] for d in DERIVED] + ["dup_flag"]
X = {name: get_column_letter(30 + i) for i, name in enumerate(DNAMES)}  # derived column letters
ALL = dict(COL, **X)
LASTCOL = X["dup_flag"]
assert X["cancel_flag"] == "AD" and LASTCOL == "BP", (X["cancel_flag"], LASTCOL)

print("Writing Clean_Data (39 derived columns, one formula per row) ...", flush=True)
ws_cl.append(list(cols) + DNAMES)
for i, r in enumerate(clean_rows):
    rr = i + 2
    ws_cl.append(r + [f(rr) for _, f, _ in DERIVED] + [dup[i]])
for row in ws_cl.iter_rows(min_row=2, max_row=LAST, min_col=DATE_COL + 1, max_col=DATE_COL + 1):
    if isinstance(row[0].value, datetime):
        row[0].number_format = "dd/mm/yyyy"
del clean_rows
for j in range(1, 30):
    h = ws_cl.cell(row=1, column=j)
    h.font, h.fill, h.alignment = F_BOLD, FILL_GREY, A_HDR
    ws_cl.column_dimensions[get_column_letter(j)].width = max(11, min(30, len(str(h.value)) + 3))
ws_cl.column_dimensions["A"].width = 17
ws_cl.column_dimensions[C["arrival_date"]].width = 17
for name, f, kind in DERIVED:
    h = ws_cl[f"{X[name]}1"]
    h.font, h.alignment = F_BOLD, A_HDR
    h.fill = FILL_AMBER if kind == "f" else FILL_TEAL
    ws_cl.column_dimensions[X[name]].width = max(12, min(24, len(name) + 3))
    h.comment = Comment(f"{'Flag' if kind == 'f' else 'Derived'} column (row formula). Row 2: {f(2)}",
                        "BUSM1042 group", width=360, height=120)
h = ws_cl[f"{X['dup_flag']}1"]
h.font, h.fill, h.alignment = F_BOLD, FILL_AMBER, A_HDR
ws_cl.column_dimensions[X["dup_flag"]].width = 11
h.comment = Comment("VALUE written by Python (DQ17): 1 if this row is identical to an EARLIER row on all 28 columns "
                    "except booking_order_id, else 0. Method: pandas df.duplicated(subset=all columns except "
                    "booking_order_id, keep='first') on the source values as stored. A 28-column duplicate test is "
                    "impractical as a formula over 119,390 rows.", "BUSM1042 group", width=360, height=150)
tbl = Table(displayName="tblClean", ref=f"A1:{LASTCOL}{LAST}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True, showColumnStripes=False,
                                    showFirstColumn=False, showLastColumn=False)
ws_cl.add_table(tbl)
ws_cl.freeze_panes = "B2"
ws_cl.sheet_properties.tabColor = "2E75B6"

# ---------------------------------------------------------------- named ranges (exact rows 2:119391)
RNG = {  # range name -> Clean_Data column name
    "rng_id": "booking_order_id", "rng_hotel": "hotel_name", "rng_is_canceled": "is_canceled",
    "rng_lead_time": "lead_time", "rng_arrival_date": "arrival_date",
    "rng_weekend_nights": "stays_in_weekend_nights", "rng_week_nights": "stays_in_week_nights",
    "rng_adults": "adults", "rng_children": "children", "rng_babies": "babies", "rng_meal": "meal",
    "rng_country": "country", "rng_segment": "market_segment", "rng_channel": "distribution_channel",
    "rng_is_repeated": "is_repeated_guest", "rng_prev_canc": "previous_cancellations",
    "rng_prev_notcanc": "previous_bookings_not_canceled", "rng_reserved_room": "reserved_room_type",
    "rng_assigned_room": "assigned_room_type", "rng_changes": "booking_changes", "rng_deposit": "deposit_type",
    "rng_agent": "agent", "rng_company": "company", "rng_waiting": "days_in_waiting_list",
    "rng_customer": "customer_type", "rng_adr": "Average_Daily_Rate", "rng_parking": "required_car_parking_spaces",
    "rng_requests": "total_of_special_requests", "rng_status": "reservation_status",
    "rng_cancel": "cancel_flag", "rng_repeat": "repeat_flag", "rng_dup": "dup_flag",
}
for name in DNAMES:
    if name not in ("cancel_flag", "repeat_flag", "dup_flag"):
        RNG[f"rng_{name}"] = name
for k, field in RNG.items():
    col = ALL[field]
    wb.defined_names[k] = DefinedName(k, attr_text=f"Clean_Data!${col}$2:${col}${LAST}")

# ================================================================ README (with the ADR_MAX named cell)
rd = ws_readme
rd.column_dimensions["A"].width = 30
rd.column_dimensions["B"].width = 112
title(rd, "A1", "BUSM1042 Assignment 3: shared, audited hotel bookings dataset (Q1–Q5)")
README_LINES = [
    ("Purpose", "One cleaned, checked copy of the supplied hotel bookings data for all five business questions "
                "(Q1–Q5) of the group report, written for an Australian hospitality association that is concerned "
                "about hotel booking cancellations among Australian tourists."),
    ("Scenario rule", "Every record is treated as a booking by an Australian tourist, whatever its recorded country "
                      "of origin. The country column is the guest's country of origin, not citizenship, so it is "
                      "NEVER used to filter or cut the data. The published source describes hotels in Portugal; for "
                      "this assignment we follow the scenario."),
    ("Source", "IBA-A3-Hotel_bookings_dataset-T32026.xlsx, sheet 'hotels' (119,390 bookings × 29 columns). "
               "Variable meanings: Antonio, de Almeida & Nunes (2019), 'Hotel booking demand datasets', Data in "
               "Brief, vol. 22, Table 1 (pp. 43–45). Data dictionary: IBA-A3-Dataset_description-T32026.xlsx."),
    ("Date built", BUILD_DATE),
    ("Before PROMPT B", "Open this file in desktop Excel, wait for the calculation to finish (Clean_Data holds about "
                        "4.5 million row formulas), then Save. This stores the calculated values. Then run PROMPT B "
                        "in Claude in Excel."),
    (None, None),
    ("SHEET MAP", None),
    ("README", "This page: purpose, scenario rule, sheet map, cleaning rules, band and flag definitions."),
    ("Raw_Data", "The untouched original: an exact copy of the source (same columns, order, rows and values, "
                 "including the texts NULL, NA and 29 February 2023). Protected without a password."),
    ("Column_Profile", "Step 1: every column profiled before cleaning (meaning, type, values, anything odd)."),
    ("Clean_Data", "Excel Table tblClean: the 29 original columns (NULL/NA texts converted to empty cells) plus 39 "
                   "derived and flag columns. Teal headers = derived, amber = flags; hover a header for its formula."),
    ("Data_Quality_Log", "The issue register DQ01–DQ20: live counts, expected counts, Match?, what was done, why, "
                         "and the effect on results."),
    ("Sensitivity", "How much each data-quality decision moves the cancellation rate, overall and in every "
                    "one-variable table."),
    ("OneVar_Summary", "Cancellation rate against one variable at a time (formula tables). PROMPT B reconciles its "
                       "PivotTables with these. No country block (scenario rule)."),
    ("Checks", "Live TRUE/FALSE checks with an ALL CHECKS PASS banner in A3."),
    (None, None),
    ("CLEANING RULES", None),
    ("No deletions", "All 119,390 bookings stay in every sheet. Problems are handled with flags, recoded copies "
                     "and sensitivity checks, never by deleting rows."),
    ("No overwrites", "Source values are never changed. Recoded values go in NEW columns next to the originals."),
    ("Text codes to blanks", "The source stores missing or not-applicable values as the text NULL (country, agent, "
                             "company) and NA (children). Clean_Data converts exactly these 129,425 cells to empty "
                             "cells; Raw_Data keeps them. Nothing else changes, and 29 February 2023 stays as text."),
    ("Cancellation", "cancel_flag = 1 when is_canceled = Yes (the paper's 1 = cancelled). It includes no-shows."),
    ("Meal (stated correction)", "meal_clean maps Undefined to SC, because the paper defines Undefined/SC as 'no meal "
                                 "package'. The original meal column is kept."),
    ("NULL = not applicable", "agent_clean and company_clean show \"None\" where the source has NULL: the paper "
                              "says NULL means not applicable, not missing (p. 45)."),
    ("Dates", "date_valid = 0 for the 96 impossible text dates; year and month are derived only for valid dates."),
    ("Duplicates (dup_flag)", "1 when a row is identical to an EARLIER row on all 28 columns except "
                              "booking_order_id, else 0. Written as a value by Python: pandas df.duplicated(subset=all "
                              "columns except booking_order_id, keep='first') on the source values as stored. Kept, "
                              "not deleted (DQ17)."),
    (None, None),
    ("BAND DEFINITIONS (inclusive)", None),
    ("lead_band", "0–7, 8–30, 31–90, 91–180, 181–365 and Over 365 (366 or more days); the same bands as Q2."),
    ("nights_band", "Total nights (weekend + week): 0, 1, 2–3, 4–7, 8+."),
    ("prev_canc_band / prev_notcanc_band", "Previous cancellations 0, 1, 2+; previous bookings not cancelled 0, 1+."),
    ("changes_band / requests_band", "0, 1, 2, 3+."),
    ("parking_band / waiting_band", "0, 1+ (parking spaces; days on the waiting list)."),
    ("adr_band", "Average_Daily_Rate: 0 or less; 0.01–50; 50.01–100; 100.01–150; 150.01–200; Over 200."),
    ("*_order columns", "1, 2, 3 ... in the natural order of each band, so tables and PivotTables sort correctly."),
    ("n < 30", "Any rate based on fewer than 30 bookings is marked unreliable."),
    (None, None),
    ("FLAG DEFINITIONS", None),
    ("flag_invalid", "Values that cannot be right: children NA, zero guests, a negative rate or the maximum rate "
                     "(DQ04, DQ10, DQ13)."),
    ("flag_undefined", "market_segment or distribution_channel = Undefined (DQ08)."),
    ("flag_outlier", "Implausible but possible: no adults with children or babies, adults 10+, children 10+, "
                     "babies 9+, parking spaces 8+ (DQ11, DQ12, DQ16)."),
    ("flag_zero_rate", "Average_Daily_Rate = 0 (DQ14)."),
    ("flag_zero_nights", "Zero total nights (DQ15)."),
    ("flag_any_quality", "flag_invalid, flag_undefined or flag_outlier: the rows excluded in the sensitivity test."),
    ("dup_flag", "Duplicate of an earlier row on all 28 variables (DQ17)."),
    (None, None),
    ("NAMED VALUES", None),
    ("ADR_MAX", "__ADR_MAX__"),
]
r = 3
for k, v in README_LINES:
    if k is None:
        r += 1
        continue
    if v is None:
        put(rd, f"A{r}", k, font=F_TTITLE)
        rd.row_dimensions[r].height = 18
    elif v == "__ADR_MAX__":
        put(rd, f"A{r}", "ADR_MAX (named cell, used by flag_invalid)", font=F_BOLD, align=Alignment(vertical="top"))
        put(rd, f"B{r}", f"=MAX({R('Average_Daily_Rate')})", fmt="#,##0.00",
            align=Alignment(horizontal="left", vertical="top"))
        wb.defined_names["ADR_MAX"] = DefinedName("ADR_MAX", attr_text=f"README!$B${r}")
    else:
        put(rd, f"A{r}", k, font=F_BOLD, align=Alignment(vertical="top", wrap_text=True))
        put(rd, f"B{r}", v, align=A_WRAP)
        rd.row_dimensions[r].height = 15 * max(1, -(-len(v) // 118))
    r += 1
presentation(rd, "808080", area=f"A1:B{r}")

# ================================================================ Column_Profile (Step 1)
pf = ws_prof
title(pf, "A1", "Column profile (Step 1): every column checked one by one, before cleaning")
put(pf, "A2", "Read from Raw_Data with every value kept as stored. Min, median, mean, max, zero counts, text counts "
              "and top-value counts are live formulas on Raw_Data; Python types and distinct counts come from the "
              "Python profile (a distinct count is impractical as a formula over 119,390 rows).", font=F_SMALL)
PF_H = 4
header_row(pf, PF_H, 1, ["#", "Column", "Meaning (Data in Brief, Table 1)", "Type", "Python types found (Python)",
                         "Distinct values (Python)", "Min", "Median", "Mean", "Max", "Zeros", "Top values (count)",
                         "Text in a numeric column", "Anything odd"], height=34)
TOPLABEL = {"is_canceled": None}


def countif_text(field, value):
    v = f'"{value}"' if isinstance(value, str) else str(value)
    return f"COUNTIF({R(field)},{v})"


ODD = {  # anything odd, as (parts) for F(); numbers are live
    "booking_order_id": ["None: one unique ID per booking (checked in Python)."],
    "deposit_type": ["Non Refund bookings were cancelled ",
                     (f'COUNTIFS({R("deposit_type")},"Non Refund",{R("is_canceled")},"Yes")/'
                      f'COUNTIF({R("deposit_type")},"Non Refund")', "0.0%"),
                     " of the time: check whether this is real behaviour or a recording practice."],
    "assigned_room_type": ["Some codes appear only as assigned types (for example I and K), never as reserved types."],
    "is_canceled": ["Coded Yes/No, while the dictionary and the paper use 0/1 (DQ01)."],
    "lead_time": ["Right-skewed: median ", (f"MEDIAN({R('lead_time')})", "0"), " but mean ",
                  (f"AVERAGE({R('lead_time')})", "0.00"), " days; ", (countif_text("lead_time", 0), "#,##0"),
                  " zero-day bookings."],
    "arrival_date": [(f"SUMPRODUCT(--({R('arrival_date')}=\"29 February 2023\"))", "#,##0"),
                     " text values \"29 February 2023\": impossible, because 2023 is not a leap year (DQ03). The "
                     "dictionary says \"Year of arrival\" but the column holds full dates (DQ02)."],
    "children": [(countif_text("children", "NA"), "#,##0"), " text values \"NA\" (DQ04); one booking with ",
                 (f"MAX({R('children')})", "0"), " children (DQ12)."],
    "adults": [(countif_text("adults", 0), "#,##0"), " bookings with no adults; maximum ",
               (f"MAX({R('adults')})", "0"), " adults in one booking (DQ11, DQ12)."],
    "babies": ["Bookings with 9 or 10 babies: ", (f"COUNTIF({R('babies')},\">=9\")", "0"), " (DQ12)."],
    "meal": [(countif_text("meal", "Undefined"), "#,##0"),
             " \"Undefined\": the paper treats Undefined/SC as no meal package (DQ09)."],
    "country": [(countif_text("country", "NULL"), "#,##0"), " text \"NULL\" (DQ05). Both \"CN\" (",
                (countif_text("country", "CN"), "#,##0"), ") and \"CHN\" (", (countif_text("country", "CHN"), "#,##0"),
                ") appear (DQ06)."],
    "market_segment": [(countif_text("market_segment", "Undefined"), "0"), " \"Undefined\" (DQ08)."],
    "distribution_channel": [(countif_text("distribution_channel", "Undefined"), "0"), " \"Undefined\" (DQ08)."],
    "is_repeated_guest": ["Coded Yes/No, while the dictionary and the paper use 0/1 (DQ01)."],
    "previous_cancellations": [(f"COUNTIFS({R('previous_cancellations')},\">=11\")", "#,##0"),
                               " bookings with 11 to 26 previous cancellations, probably repeated group records."],
    "agent": [(countif_text("agent", "NULL"), "#,##0"), " text \"NULL\" = no agent (not applicable, DQ07). IDs are "
              "stored as numbers but are codes."],
    "company": [(countif_text("company", "NULL"), "#,##0"), " text \"NULL\" = no company (not applicable, DQ07). "
                "IDs are stored as numbers but are codes."],
    "Average_Daily_Rate": ["Negative rate ", (f"MIN({R('Average_Daily_Rate')})", "0.00"), " and maximum ",
                           (f"MAX({R('Average_Daily_Rate')})", "#,##0"), " (DQ13); ",
                           (countif_text("Average_Daily_Rate", 0), "#,##0"), " zero rates (DQ14)."],
    "required_car_parking_spaces": [(countif_text("required_car_parking_spaces", 8), "0"),
                                    " bookings need 8 spaces (DQ16); no booking with a space was cancelled (DQ20)."],
    "reserved_room_type": ["Reserved and assigned types differ in ",
                           (f"SUMPRODUCT(--({R('reserved_room_type')}<>{R('assigned_room_type')}))", "#,##0"),
                           " bookings (DQ20)."],
    "reservation_status": ["Canceled + No-Show = ", (f"COUNTIF({R('reservation_status')},\"<>Check-Out\")", "#,##0"),
                           ", the same as is_canceled = Yes (DQ18)."],
    "stays_in_weekend_nights": ["Zero total nights (weekend + week) in ",
                                (f"SUMPRODUCT(--(({R('stays_in_weekend_nights')}+{R('stays_in_week_nights')})=0))",
                                 "#,##0"), " bookings (DQ15)."],
}
ODD_DEFAULT = "Nothing odd found."
for i, p in enumerate(profile):
    rr = PF_H + 1 + i
    name = p["name"]
    rng_ = R(name)
    band = i % 2 == 1
    body_cell(pf, rr, 1, i + 1, "0", band, align=A_CTR)
    body_cell(pf, rr, 2, name, None, band, bold=True)
    body_cell(pf, rr, 3, MEANING[name], None, band, wrap=True)
    kind = "Measure (numeric)" if p["measure"] else ("ID" if name in ("booking_order_id",) else
                                                     "Dimension (date)" if name == "arrival_date" else
                                                     "Dimension (ID code)" if name in ("agent", "company") else
                                                     "Dimension (categorical)")
    body_cell(pf, rr, 4, kind, None, band, wrap=True)
    body_cell(pf, rr, 5, p["types"], None, band, wrap=True)
    body_cell(pf, rr, 6, p["distinct"], FMT_N, band)
    if p["measure"] or name == "arrival_date":
        dfmt = "dd/mm/yyyy" if name == "arrival_date" else ("#,##0.00" if name == "Average_Daily_Rate" else "#,##0")
        body_cell(pf, rr, 7, f"=MIN({rng_})", dfmt, band)
        body_cell(pf, rr, 8, f"=MEDIAN({rng_})" if p["measure"] else "n/a",
                  ("#,##0.000" if name == "Average_Daily_Rate" else "#,##0.0") if p["measure"] else None, band)
        body_cell(pf, rr, 9, f"=AVERAGE({rng_})" if p["measure"] else "n/a", "#,##0.00" if p["measure"] else None,
                  band)
        body_cell(pf, rr, 10, f"=MAX({rng_})", dfmt, band)
        body_cell(pf, rr, 11, f"=COUNTIF({rng_},0)" if p["measure"] else "n/a", FMT_N if p["measure"] else None,
                  band)
        body_cell(pf, rr, 13, f"=SUMPRODUCT(--ISTEXT({rng_}))", FMT_N, band)
    else:
        for j in (7, 8, 9, 10, 11, 13):
            body_cell(pf, rr, j, "n/a", None, band, align=A_CTR)
    # top values with live counts
    if name == "booking_order_id":
        top = "All values unique (one per booking)."
    elif name == "arrival_date":
        top = F("Real dates: ", (f"SUMPRODUCT(--ISNUMBER({rng_}))", "#,##0"), "; text: ",
                (f"SUMPRODUCT(--ISTEXT({rng_}))", "#,##0"))
    else:
        parts = []
        for j, val in enumerate(p["top"]):
            label = str(val)
            crit = f'"{val}"' if isinstance(val, str) else repr(val)
            parts += [("; " if j else "") + f"{label} (", (f"COUNTIF({rng_},{crit})", "#,##0"), ")"]
        if p["distinct"] > 5:
            parts.append("; …")
        top = F(*parts)
    body_cell(pf, rr, 12, top, None, band, wrap=True)
    odd = ODD.get(name)
    body_cell(pf, rr, 14, F(*odd) if odd and any(not isinstance(x, str) for x in odd) else (odd[0] if odd else
                                                                                           ODD_DEFAULT), None, band,
              wrap=True)
    pf.row_dimensions[rr].height = 62
for col_letter, w in zip("ABCDEFGHIJKLMN", [4, 26, 46, 16, 20, 11, 11, 10, 10, 11, 9, 46, 11, 46]):
    pf.column_dimensions[col_letter].width = w
PF_END = PF_H + len(profile)
put(pf, f"A{PF_END + 1}", "Measures = 13 numeric columns; dimensions = categories, dates and ID codes. Meanings "
                          "from Antonio, de Almeida & Nunes (2019, Table 1, pp. 43–45).", font=F_NOTE)
pf.freeze_panes = f"C{PF_H + 1}"
presentation(pf, "70AD47", area=f"A1:N{PF_END + 1}")

# ================================================================ OneVar_Summary
ov = ws_ov
title(ov, "A1", "Cancellation rate against one variable at a time")
put(ov, "A2", "Formula tables on tblClean (cancellations include no-shows). Rates based on fewer than 30 bookings "
              "are marked unreliable. There is no country block: country is never used to cut the data "
              "(scenario rule).", font=F_SMALL)
put(ov, "A3", "All bookings", font=F_BOLD)
put(ov, "B3", "=COUNTA(rng_id)", fmt=FMT_N, font=F_BOLD)
put(ov, "C3", "=SUM(rng_cancel)", fmt=FMT_N, font=F_BOLD)
put(ov, "D3", "=C3/B3", fmt=FMT_P, font=F_BOLD)
put(ov, "E3", "bookings / cancelled / overall rate", font=F_SMALL)
wb.defined_names["N_ALL"] = DefinedName("N_ALL", attr_text="OneVar_Summary!$B$3")
wb.defined_names["C_ALL"] = DefinedName("C_ALL", attr_text="OneVar_Summary!$C$3")


def counts_desc(field_values):
    return [k for k, _ in Counter(field_values).most_common()]


def col_values(name):
    return df[name].tolist()


MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def q(v):
    return f'"{v}"' if isinstance(v, str) else str(v)


def ordered(rng, labels):
    return [(lab, [(rng, i + 1)]) for i, lab in enumerate(labels)]


def by_text(rng, values):
    return [(v, [(rng, v)]) for v in values]


with_children_vals = Counter("Yes" if ((0 if c_ == "NA" else c_) + b_) > 0 else "No"
                             for c_, b_ in zip(df["children"], df["babies"]))
room_vals = Counter("Same" if a == b else "Different" for a, b in zip(df["reserved_room_type"], df["assigned_room_type"]))
meal_clean_vals = Counter("SC" if m == "Undefined" else m for m in df["meal"])
BLOCKS = [  # key, title, categories [(label, [(range, value)])], caution, note
    ("hotel_name", "hotel_name", by_text("rng_hotel", counts_desc(col_values("hotel_name"))), False, None),
    ("lead_band", "lead_band (days before arrival)",
     ordered("rng_lead_band_order", ["0–7", "8–30", "31–90", "91–180", "181–365", "Over 365"]), False, None),
    ("deposit_type", "deposit_type", by_text("rng_deposit", counts_desc(col_values("deposit_type"))), False, None),
    ("market_segment", "market_segment", by_text("rng_segment", counts_desc(col_values("market_segment"))), False,
     None),
    ("distribution_channel", "distribution_channel",
     by_text("rng_channel", counts_desc(col_values("distribution_channel"))), False, None),
    ("customer_type", "customer_type", by_text("rng_customer", counts_desc(col_values("customer_type"))), False,
     None),
    ("is_repeated_guest", "is_repeated_guest", by_text("rng_is_repeated", counts_desc(col_values("is_repeated_guest"))),
     False, None),
    ("prev_canc_band", "prev_canc_band (previous cancellations)", ordered("rng_prev_canc_band_order", ["0", "1", "2+"]),
     False, None),
    ("prev_notcanc_band", "prev_notcanc_band (previous bookings not cancelled)",
     ordered("rng_prev_notcanc_band_order", ["0", "1+"]), False, None),
    ("meal_clean", "meal_clean (Undefined counted as SC)",
     by_text("rng_meal_clean", [k for k, _ in meal_clean_vals.most_common()]), False, None),
    ("changes_band", "changes_band (booking changes)", ordered("rng_changes_band_order", ["0", "1", "2", "3+"]), False,
     None),
    ("requests_band", "requests_band (special requests)", ordered("rng_requests_band_order", ["0", "1", "2", "3+"]),
     False, None),
    ("waiting_band", "waiting_band (days on the waiting list)", ordered("rng_waiting_band_order", ["0", "1+"]), False,
     None),
    ("nights_band", "nights_band (total nights)", ordered("rng_nights_band_order", ["0", "1", "2–3", "4–7", "8+"]),
     False, None),
    ("adr_band", "adr_band (Average_Daily_Rate)",
     ordered("rng_adr_band_order", ["0 or less", "0.01–50", "50.01–100", "100.01–150", "150.01–200", "Over 200"]),
     False, None),
    ("with_children", "with_children (children or babies)", by_text("rng_with_children",
                                                                     [k for k, _ in with_children_vals.most_common()]),
     False, None),
    ("arrival_month", "arrival_month (valid dates only; the 96 invalid dates are shown separately)",
     [(m, [("rng_arrival_month", i + 1)]) for i, m in enumerate(MONTHS)] + [("Invalid date", [("rng_date_valid", 0)])],
     False, "month"),
    ("arrival_year", "arrival_year (2022 and 2024 are part-years, see DQ19)",
     [(str(y), [("rng_arrival_year", y)]) for y in (2022, 2023, 2024)] + [("Invalid date", [("rng_date_valid", 0)])],
     False, "year"),
    ("parking_band", f"parking_band ({CAUTION})", ordered("rng_parking_band_order", ["0", "1+"]), True, None),
    ("room_changed", f"room_changed: reserved vs assigned room type ({CAUTION})",
     by_text("rng_room_changed", [k for k, _ in room_vals.most_common()]), True, None),
    ("reserved_room_type", f"reserved_room_type ({CAUTION}); room-type codes in letter order",
     by_text("rng_reserved_room", sorted(set(col_values("reserved_room_type")))), True, None),
]
assert len(BLOCKS) == 21


def crit(cs):
    return ",".join(f"{r_},{q(v)}" for r_, v in cs)


OV_ROWS = {}  # block key -> dict(first=, last=, total=, cats=[(label, row, crit)])
r = 6
for bi, (key, btitle, cats, caution, special) in enumerate(BLOCKS):
    ttitle(ov, f"A{r}", f"Block {bi + 1}: {btitle}")
    if caution:
        ov[f"A{r}"].font = Font(name=FONT, size=12, bold=True, color="C00000")
    header_row(ov, r + 1, 1, ["Category", "Bookings", "Cancelled", "Cancellation rate", "Share of all bookings",
                              "Share of all cancellations", "n < 30: unreliable"], height=32)
    hc = ov.cell(row=r + 1, column=8, value="helper: rate if n ≥ 30")
    hc.font, hc.alignment = F_HELP, A_HDR
    first = r + 2
    rows_ = []
    for i, (lab, cs) in enumerate(cats):
        rr = first + i
        band = i % 2 == 1
        body_cell(ov, rr, 1, lab, "@", band)
        body_cell(ov, rr, 2, f"=COUNTIFS({crit(cs)})", FMT_N, band)
        body_cell(ov, rr, 3, f"=COUNTIFS({crit(cs)},rng_cancel,1)", FMT_N, band)
        body_cell(ov, rr, 4, f'=IF(B{rr}=0,"n/a",C{rr}/B{rr})', FMT_P, band)
        body_cell(ov, rr, 5, f"=B{rr}/N_ALL", FMT_P, band)
        body_cell(ov, rr, 6, f"=C{rr}/C_ALL", FMT_P, band)
        body_cell(ov, rr, 7, f'=IF(B{rr}<30,"n < 30: unreliable","")', None, band, align=A_CTR)
        helper = '""' if lab == "Invalid date" else f'IF(B{rr}>=30,D{rr},"")'
        hc = ov.cell(row=rr, column=8, value=f"={helper}")
        hc.font, hc.number_format = F_HELP, FMT_P
        rows_.append((lab, rr, cs))
    last = first + len(cats) - 1
    tot = last + 1
    body_cell(ov, tot, 1, "Total", None, total=True)
    body_cell(ov, tot, 2, f"=SUM(B{first}:B{last})", FMT_N, total=True)
    body_cell(ov, tot, 3, f"=SUM(C{first}:C{last})", FMT_N, total=True)
    body_cell(ov, tot, 4, f"=C{tot}/B{tot}", FMT_P, total=True)
    body_cell(ov, tot, 5, f"=SUM(E{first}:E{last})", FMT_P, total=True)
    body_cell(ov, tot, 6, f"=SUM(F{first}:F{last})", FMT_P, total=True)
    body_cell(ov, tot, 7, "", None, total=True)
    H = f"H{first}:H{last}"
    A_ = f"A{first}:A{last}"
    small = f'COUNTIF(B{first}:B{last},"<30")'
    sentence = F("What stands out: ", ("=", f"INDEX({A_},MATCH(MAX({H}),{H},0))"), " has the highest cancellation rate (",
                 (f"MAX({H})", "0.0%"), ") and ", ("=", f"INDEX({A_},MATCH(MIN({H}),{H},0))"), " the lowest (",
                 (f"MIN({H})", "0.0%"), "), a gap of ", (f"(MAX({H})-MIN({H}))*100", "0.0"), " percentage points.",
                 ("=", f'IF({small}=0,"",IF({small}=1," One category has"," "&{small}&" categories have")&'
                       f'" fewer than 30 bookings: unreliable, left out of this comparison.")'),
                 ("=", '" Valid dates only; the 96 invalid dates are shown on their own row."' if special == "month"
                  else '" 2022 and 2024 are part-years, so compare rates, not totals."' if special == "year"
                  else '" Use with caution: this field may be recorded at check-in, so it may reflect the outcome."'
                  if caution else '""'))
    c = ov.cell(row=tot + 1, column=1, value=sentence)
    c.font = Font(name=FONT, size=10, italic=True, color="404040")
    OV_ROWS[key] = dict(first=first, last=last, total=tot, cats=rows_, title=btitle, caution=caution,
                        stands=tot + 1)
    r = tot + 4
OV_END = r
for col_letter, w in zip("ABCDEFGH", [26, 12, 12, 13, 13, 14, 17, 11]):
    ov.column_dimensions[col_letter].width = w
ov.freeze_panes = "A5"
presentation(ov, "70AD47", landscape=False, area=f"A1:G{OV_END}")

# ================================================================ Sensitivity
se = ws_sens
title(se, "A1", "Sensitivity: how much could each data-quality decision change the results?")
put(se, "A2", "All values are formulas on tblClean. Changes are in percentage points (points). 'Expected' values come "
              "from the brief and are typed; Match? compares them after rounding.", font=F_SMALL)
ttitle(se, "A4", "Part 1: overall cancellation rate under each decision")
S1_H = 5
header_row(se, S1_H, 1, ["Test", "Bookings used", "Cancelled", "Overall cancellation rate",
                         "Change vs baseline (points)", "Expected (brief)", "Match?"], height=32)
TESTS = [
    ("base", "All bookings (baseline)", "=COUNTA(rng_id)", "=SUM(rng_cancel)", "=AVERAGE(rng_cancel)", 0.3704),
    ("any", "Excluding flag_any_quality = 1 (invalid, undefined or outlier rows)", "=COUNTIF(rng_flag_any_quality,0)",
     "=COUNTIFS(rng_flag_any_quality,0,rng_cancel,1)", "=AVERAGEIFS(rng_cancel,rng_flag_any_quality,0)", 0.3707),
    ("invund", "Excluding flag_invalid and flag_undefined only", "=COUNTIFS(rng_flag_invalid,0,rng_flag_undefined,0)",
     "=COUNTIFS(rng_flag_invalid,0,rng_flag_undefined,0,rng_cancel,1)",
     "=AVERAGEIFS(rng_cancel,rng_flag_invalid,0,rng_flag_undefined,0)", 0.3707),
    ("zn", "Excluding zero-night stays", "=COUNTIF(rng_flag_zero_nights,0)",
     "=COUNTIFS(rng_flag_zero_nights,0,rng_cancel,1)", "=AVERAGEIFS(rng_cancel,rng_flag_zero_nights,0)", 0.3724),
    ("dup", "Excluding duplicates (dup_flag = 1)", "=COUNTIF(rng_dup,0)", "=COUNTIFS(rng_dup,0,rng_cancel,1)",
     "=AVERAGEIFS(rng_cancel,rng_dup,0)", 0.2729),
    ("noshow", "Excluding no-shows (only reservation_status = Canceled counts as cancelled; all bookings kept)",
     "=COUNTA(rng_id)", '=COUNTIF(rng_status,"Canceled")', "=C{r}/B{r}", 0.3603),
]
S1 = {}
for i, (key, lab, nf, cf, rf, exp) in enumerate(TESTS):
    rr = S1_H + 1 + i
    band = i % 2 == 1
    S1[key] = rr
    body_cell(se, rr, 1, lab, None, band, wrap=True)
    body_cell(se, rr, 2, nf, FMT_N, band)
    body_cell(se, rr, 3, cf, FMT_N, band)
    body_cell(se, rr, 4, rf.format(r=rr), FMT_P2, band)
    body_cell(se, rr, 5, f"=(D{rr}-$D${S1_H + 1})*100", FMT_PT, band)
    c = body_cell(se, rr, 6, exp, FMT_P2, band)
    c.font = F_EXP
    body_cell(se, rr, 7, f'=IF(ROUND(D{rr},4)=F{rr},"{OK}","{BAD}")', None, band, align=A_CTR)
    se.row_dimensions[rr].height = 30
S1_END = S1_H + len(TESTS)
match_rules(se, f"G{S1_H + 1}:G{S1_END}")
put(se, f"A{S1_END + 1}", "The two decisions that matter are duplicates and no-shows; every other flag moves the overall "
                          "rate by well under 0.1 points.", font=F_NOTE)
BASE = f"Sensitivity!$D${S1['base']}"

# Part 3 first (detail), so that Part 2 can point at it
EXPECTED_LARGEST = {"hotel_name": 0.07, "lead_band": 0.05, "deposit_type": 0.01, "distribution_channel": 0.03,
                    "is_repeated_guest": 0.18, "meal_clean": 0.15, "prev_canc_band": 0.24, "market_segment": 0.90,
                    "customer_type": 2.19, "nights_band": 1.33, "adr_band": 1.21, "reserved_room_type": 1.80}
S2_T = S1_END + 3
S2_H = S2_T + 1
S2_FIRST = S2_H + 1
S2_END = S2_FIRST + len(BLOCKS) - 1
S3_T = S2_END + 4
ttitle(se, f"A{S3_T}", "Part 3: every one-variable table, baseline vs excluding flag_any_quality = 1")
S3_H = S3_T + 1
header_row(se, S3_H, 1, ["Variable", "Category", "Bookings", "Cancelled", "Baseline rate", "Flagged rows",
                         "Flagged: invalid", "Flagged: outlier", "Flagged: undefined", "Flagged and cancelled",
                         "Bookings after excluding", "Rate after excluding", "Change (points)",
                         "Counted (n ≥ 30): |change|"], height=46)
S3 = {}
rr = S3_H + 1
for bi, (key, btitle, cats, caution, special) in enumerate(BLOCKS):
    info = OV_ROWS[key]
    first = rr
    for i, (lab, ovr, cs) in enumerate(info["cats"]):
        band = bi % 2 == 1
        cr = crit(cs)
        body_cell(se, rr, 1, key, None, band)
        body_cell(se, rr, 2, lab, "@", band)
        body_cell(se, rr, 3, f"=OneVar_Summary!B{ovr}", FMT_N, band)
        body_cell(se, rr, 4, f"=OneVar_Summary!C{ovr}", FMT_N, band)
        body_cell(se, rr, 5, f"=OneVar_Summary!D{ovr}", FMT_P2, band)
        body_cell(se, rr, 6, f"=COUNTIFS({cr},rng_flag_any_quality,1)", FMT_N, band)
        body_cell(se, rr, 7, f"=COUNTIFS({cr},rng_flag_invalid,1)", FMT_N, band)
        body_cell(se, rr, 8, f"=COUNTIFS({cr},rng_flag_outlier,1)", FMT_N, band)
        body_cell(se, rr, 9, f"=COUNTIFS({cr},rng_flag_undefined,1)", FMT_N, band)
        body_cell(se, rr, 10, f"=COUNTIFS({cr},rng_flag_any_quality,1,rng_cancel,1)", FMT_N, band)
        body_cell(se, rr, 11, f"=C{rr}-F{rr}", FMT_N, band)
        body_cell(se, rr, 12, f'=IF(K{rr}=0,"n/a",(D{rr}-J{rr})/K{rr})', FMT_P2, band)
        body_cell(se, rr, 13, f'=IF(ISNUMBER(L{rr}),(L{rr}-E{rr})*100,"n/a")', FMT_PT, band)
        body_cell(se, rr, 14, f'=IF(AND(C{rr}>=30,ISNUMBER(M{rr})),ABS(M{rr}),"")', "0.0000", band)
        rr += 1
    S3[key] = (first, rr - 1)
S3_END = rr - 1

ttitle(se, f"A{S2_T}", "Part 2: largest change in any category with 30 or more bookings, per one-variable table")
header_row(se, S2_H, 1, ["Variable", "Categories with 30+ bookings", "Largest change (points)", "Where",
                         "Signed change (points)", "Expected (brief)", "Match?", "Why (one sentence)"], height=32)
for i, (key, btitle, cats, caution, special) in enumerate(BLOCKS):
    rr = S2_FIRST + i
    band = i % 2 == 1
    a, b = S3[key]
    Ncol, Bcol, Ccol, Fcol = f"N{a}:N{b}", f"B{a}:B{b}", f"C{a}:C{b}", f"F{a}:F{b}"
    pos = f"MATCH(C{rr},{Ncol},0)"
    body_cell(se, rr, 1, key, None, band)
    body_cell(se, rr, 2, f'=COUNTIF({Ccol},">=30")', "0", band)
    body_cell(se, rr, 3, f"=MAX({Ncol})", "0.00", band)
    body_cell(se, rr, 4, f"=INDEX({Bcol},{pos})", None, band)
    body_cell(se, rr, 5, f"=INDEX(M{a}:M{b},{pos})", FMT_PT, band)
    exp = EXPECTED_LARGEST.get(key)
    c = body_cell(se, rr, 6, exp if exp is not None else "n/a", "0.00" if exp is not None else None, band)
    c.font = F_EXP
    body_cell(se, rr, 7, f'=IF(ISNUMBER(F{rr}),IF(ROUND(C{rr},2)=F{rr},"{OK}","{BAD}"),"no brief value")', None, band,
              align=A_CTR)
    n_, f_, inv_, out_, fc_, c_ = (f"INDEX({x},{pos})" for x in (Ccol, Fcol, f"G{a}:G{b}", f"H{a}:H{b}",
                                                                f"J{a}:J{b}", f"D{a}:D{b}"))
    why = F(("=", f'IF(G{rr}="{BAD}","The brief lists "&TEXT(F{rr},"0.00")&", but the exact change is "&'
                  f'TEXT(C{rr},"0.0000")&" points, which rounds to "&TEXT(C{rr},"0.00")&"; the computed value is kept '
                  f'and the brief value is not changed. ","")'),
            ("=", f'IF(C{rr}<0.5,"Small: under 0.5 points, so the flagged rows hardly matter here.",'
                  f'"In "&A{rr}&", category \'"&D{rr}&"\' has only "&TEXT({n_},"#,##0")&" bookings, and "&TEXT({f_},"#,##0")&" of them ("&'
                  f'TEXT({f_}/{n_},"0.0%")&") are flagged, mostly as "&IF({inv_}>={out_},"invalid values (zero guests, '
                  f'missing children or an impossible rate)","outliers (no adults, extreme party sizes or parking)")&'
                  f'"; they cancel at "&TEXT({fc_}/{f_},"0.0%")&" against "&TEXT({c_}/{n_},"0.0%")&" for the whole '
                  f'category, so teammates using this variable must report it.")'))
    body_cell(se, rr, 8, why, None, band, wrap=True)
    se.row_dimensions[rr].height = 48
match_rules(se, f"G{S2_FIRST}:G{S2_END}")
put(se, f"A{S2_END + 1}", "Flagged rows (no adults, zero guests, extreme party sizes and prices) are concentrated in a "
                          "few small categories, which is why a handful of tables move by more than 0.5 points.",
    font=F_NOTE)
SENS_LARGEST = {key: f"Sensitivity!$C${S2_FIRST + i}" for i, (key, *_r) in enumerate(BLOCKS)}

# Part 4: other values quoted in the Data_Quality_Log and the Word notes (all live)
S4_T = S3_END + 3
ttitle(se, f"A{S4_T}", "Part 4: other values quoted in the Data_Quality_Log and the Data Quality Notes")
S4_H = S4_T + 1
header_row(se, S4_H, 1, ["Value", "Live value", "Formula basis"], height=22)
SUPP = [  # key, label, formula, fmt
    ("und_rate", "Overall rate excluding flag_undefined = 1 (DQ08)", "=AVERAGEIFS(rng_cancel,rng_flag_undefined,0)",
     "0.0000%"),
    ("und_chg", "Change vs baseline, excluding flag_undefined (points)", f"=(B{{und_rate}}-{BASE})*100", "+0.0000;-0.0000"),
    ("sc_raw_n", "meal = SC in the original column (bookings)", '=COUNTIF(rng_meal,"SC")', FMT_N),
    ("sc_raw_rate", "meal = SC in the original column (rate)", '=AVERAGEIFS(rng_cancel,rng_meal,"SC")', FMT_P),
    ("sc_clean_n", "meal_clean = SC (bookings)", '=COUNTIF(rng_meal_clean,"SC")', FMT_N),
    ("sc_clean_rate", "meal_clean = SC (rate)", '=AVERAGEIFS(rng_cancel,rng_meal_clean,"SC")', FMT_P),
    ("zr_comp", "Zero-rate bookings that are Complementary (DQ14)",
     '=COUNTIFS(rng_flag_zero_rate,1,rng_segment,"Complementary")', FMT_N),
    ("zn_zr", "Zero-night stays that also have a rate of 0 (DQ15)", "=COUNTIFS(rng_flag_zero_nights,1,rng_flag_zero_rate,1)",
     FMT_N),
    ("adr_median", "Median Average_Daily_Rate", "=MEDIAN(rng_adr)", "0.000"),
    ("adr_ratio", "Maximum rate ÷ median rate (DQ13)", f"=ADR_MAX/B{{adr_median}}", "0.0"),
    ("dup_go", "Duplicates that are Groups or Offline TA/TO (DQ17)",
     '=COUNTIFS(rng_dup,1,rng_segment,"Groups")+COUNTIFS(rng_dup,1,rng_segment,"Offline TA/TO")', FMT_N),
    ("dup_go_sh", "…as a share of all duplicates", "=B{dup_go}/SUM(rng_dup)", FMT_P),
    ("go_sh", "Groups + Offline TA/TO as a share of all bookings",
     '=(COUNTIF(rng_segment,"Groups")+COUNTIF(rng_segment,"Offline TA/TO"))/N_ALL', FMT_P),
    ("over365_sh", "Over 365 band: share of all bookings (Q2)", "=COUNTIF(rng_lead_band_order,6)/N_ALL", FMT_P),
    ("over365_nr", "Over 365 band: share that is Non Refund (Q2)",
     '=COUNTIFS(rng_lead_band_order,6,rng_deposit,"Non Refund")/COUNTIF(rng_lead_band_order,6)', FMT_P),
    ("pc11", "Bookings with 11 to 26 previous cancellations (Q5)", '=COUNTIFS(rng_prev_canc,">=11",rng_prev_canc,"<=26")',
     FMT_N),
    ("pc11_go", "…of which Groups or Offline TA/TO",
     '=COUNTIFS(rng_prev_canc,">=11",rng_segment,"Groups")+COUNTIFS(rng_prev_canc,">=11",rng_segment,"Offline TA/TO")',
     FMT_N),
    ("pc11_go_sh", "…as a share of those bookings", "=B{pc11_go}/B{pc11}", FMT_P),
    ("park_n", "Bookings with 1+ parking spaces (DQ20)", '=COUNTIF(rng_parking,">=1")', FMT_N),
    ("park_c", "…of which cancelled", '=COUNTIFS(rng_parking,">=1",rng_cancel,1)', FMT_N),
    ("rc_n", "Bookings with a changed room type (DQ20)", '=COUNTIF(rng_room_changed,"Different")', FMT_N),
    ("rc_rate", "…their cancellation rate", '=AVERAGEIFS(rng_cancel,rng_room_changed,"Different")', FMT_P),
    ("date_min", "Earliest valid arrival date", "=MIN(rng_arrival_date)", "d mmmm yyyy"),
    ("date_max", "Latest valid arrival date", "=MAX(rng_arrival_date)", "d mmmm yyyy"),
    ("feb_valid", "February arrivals with a valid date", "=COUNTIF(rng_arrival_month,2)", FMT_N),
]
S4 = {}
for i, (key, *_r) in enumerate(SUPP):
    S4[key] = S4_H + 1 + i
for i, (key, lab, fml, fmt) in enumerate(SUPP):
    rr = S4[key]
    band = i % 2 == 1
    body_cell(se, rr, 1, lab, None, band, wrap=True)
    body_cell(se, rr, 2, fml.format(**S4), fmt, band)
    body_cell(se, rr, 3, fml.format(**S4).lstrip("="), None, band)
    se[f"C{rr}"].font = F_HELP
S4_END = S4_H + len(SUPP)
SV = {k: f"Sensitivity!$B${v}" for k, v in S4.items()}
for col_letter, w in zip("ABCDEFGHIJKLMN", [44, 14, 14, 14, 14, 13, 12, 12, 12, 13, 13, 13, 12, 14]):
    se.column_dimensions[col_letter].width = w
se.column_dimensions["H"].width = 70
presentation(se, "70AD47", area=f"A1:N{S4_END}")

# ================================================================ Data_Quality_Log (the issue register)
dq = ws_dq
title(dq, "A1", "Data quality log: issues DQ01–DQ20, what was done, why, and the effect on results")
put(dq, "A2", "Rows affected are live COUNTIF/COUNTIFS formulas on Raw_Data or tblClean; numbers inside the text are "
              "built with TEXT() from live cells. Quality dimensions follow Vidgen, Kirshner & Tan (2019, pp. 57–58).",
    font=F_SMALL)
DQ_H = 4
header_row(dq, DQ_H, 1, ["ID", "Quality dimension", "Column(s)", "What was found", "Rows affected (live formula)",
                         "Expected", "Match?", "What was done", "Why", "Effect on results",
                         "Questions affected (Q1–Q5)"], height=34)


def T(expr, fmt="#,##0"):
    return (expr, fmt)


RAWC, RAWO = R("is_canceled"), R("is_repeated_guest")
ANYN = "COUNTIF(rng_flag_any_quality,1)"
ANYCHG = f"Sensitivity!$E${S1['any']}"
SEE_A5 = ["See Sensitivity sheet: excluding all ", T(ANYN), " flagged rows changes the overall rate by ",
          T(ANYCHG, "+0.00;-0.00"), " points"]
MINADR = R("Average_Daily_Rate")
DQ = [  # id, dimension, columns, found(parts), rows formula, expected, done, why(parts), effect(parts), questions
    ("DQ01", "Consistency", "is_canceled, is_repeated_guest",
     ["is_canceled and is_repeated_guest are coded Yes/No; the dictionary and the paper say 0/1"],
     f'=SUMPRODUCT((({RAWC}="Yes")+({RAWC}="No"))*(({RAWO}="Yes")+({RAWO}="No")))', 119390,
     "Recoded to 1/0 in new flag columns (cancel_flag, repeat_flag)", ["The average of a 1/0 flag is the rate"],
     ["None; coding only"], "Q1–Q5"),
    ("DQ02", "Consistency", "arrival_date", ['arrival_date holds full dates, but the dictionary says "Year of arrival"'],
     f"=COUNTA({R('arrival_date')})", 119390, "Year and month derived from the date (arrival_year, arrival_month)",
     ["Lets the data be analysed by month"], ["None"], "Q1"),
    ("DQ03", "Validity", "arrival_date",
     [T(f'SUMPRODUCT(--({R("arrival_date")}="29 February 2023"))'),
      ' rows hold the impossible text date "29 February 2023" (most likely the original 29 February 2016 after the '
      'years were shifted)'],
     f'=SUMPRODUCT(--({R("arrival_date")}="29 February 2023"))', 96,
     "Flagged (date_valid = 0); left out of date-based analysis only",
     ["The true date cannot be checked, so no new date is guessed"],
     ["February arrivals are undercounted by ", T("COUNTIF(rng_date_valid,0)"), "; no other result changes"], "Q1"),
    ("DQ04", "Completeness", "children", ['children stored as the text "NA"'], f'=COUNTIF({R("children")},"NA")', 4,
     "Converted to a blank; flagged invalid (flag_invalid)", ["The value cannot be checked"],
     ["Included in the sensitivity test. "] + SEE_A5, "All (flag test only)"),
    ("DQ05", "Completeness", "country", ['country stored as the text "NULL"'], f'=COUNTIF({R("country")},"NULL")', 488,
     'Converted to a blank, labelled "Unknown" (country_clean)',
     ["Country is never used to filter (scenario rule)"], ["None"], "None (country not used)"),
    ("DQ06", "Consistency", "country",
     ['Both "CN" (', T(f'COUNTIF({R("country")},"CN")'), ') and "CHN" (', T(f'COUNTIF({R("country")},"CHN")'),
      ") appear; the paper uses 3-letter ISO codes"], f'=COUNTIF({R("country")},"CN")', 1279,
     "Flagged only, NOT merged",
     ["It is probably the same country, but this cannot be verified, and country is not used"], ["None"],
     "None (country not used)"),
    ("DQ07", "Completeness", "agent, company",
     ['agent "NULL" (', T(f'COUNTIF({R("agent")},"NULL")'), ') and company "NULL" (',
      T(f'COUNTIF({R("company")},"NULL")'), ")"],
     F(T(f'COUNTIF({R("agent")},"NULL")'), " / ", T(f'COUNTIF({R("company")},"NULL")')), "16,340 / 112,593",
     'Converted to blanks and recoded "None" (agent_clean, company_clean)',
     ["The paper says NULL means not applicable (p. 45); an optional field left empty is still complete (Vidgen, "
      "Kirshner & Tan 2019, p. 57)"], ["None"], "None"),
    ("DQ08", "Completeness", "market_segment, distribution_channel",
     ['market_segment "Undefined" (', T('COUNTIF(rng_segment,"Undefined")'), ') or distribution_channel "Undefined" (',
      T('COUNTIF(rng_channel,"Undefined")'), ")"], "=COUNTIF(rng_flag_undefined,1)", 5,
     "Flagged (flag_undefined), kept", ["The true value cannot be checked"],
     [("=", f'IF(ABS({SV["und_chg"]})<0.01,"Excluding them moves the overall rate by less than 0.01 points",'
            f'"Excluding them moves the overall rate by "&TEXT({SV["und_chg"]},"+0.00;-0.00")&" points")')],
     "Q3, Q4"),
    ("DQ09", "Consistency", "meal",
     ['meal "Undefined" (', T(f'COUNTIF({R("meal")},"Undefined")'),
      "); the paper defines Undefined/SC as no meal package"], f'=COUNTIF({R("meal")},"Undefined")', 1169,
     "meal_clean maps Undefined → SC (stated assumption); the original meal column is kept",
     ["The paper documents the meaning (Table 1, p. 44)"],
     ["SC rate goes from ", T(SV["sc_raw_rate"], "0.0%"), " (", T(SV["sc_raw_n"]), " bookings) to ",
      T(SV["sc_clean_rate"], "0.0%"), " (", T(SV["sc_clean_n"]), ")"], "None (meal not used)"),
    ("DQ10", "Validity", "adults, children, babies", ["No guests at all (adults + children + babies = 0)"],
     "=COUNTIF(rng_guests_total,0)", 180, "Flagged invalid (flag_invalid), kept",
     ["A booking with no guests cannot be right"], SEE_A5, "All (flag test only)"),
    ("DQ11", "Validity", "adults, children, babies", ["No adults but children or babies present"],
     '=COUNTIFS(rng_adults,0,rng_guests_total,">0")', 223, "Flagged outlier (flag_outlier), kept",
     ["Unusual, but possible (for example, a booking error or a minors' group)"], SEE_A5,
     "All (flag test only)"),
    ("DQ12", "Validity", "adults, children, babies",
     ["Extreme party sizes: adults ≥ 10 (", T('COUNTIF(rng_adults,">=10")'), " rows), children 10 (",
      T('COUNTIF(rng_children,">=10")'), "), babies 9 or 10 (", T('COUNTIF(rng_babies,">=9")'), ")"],
     '=SUMPRODUCT(--(((rng_adults>=10)+(rng_children>=10)+(rng_babies>=9))>0))', 16,
     "Flagged outlier (flag_outlier), kept", ["Possibly group entries or typing errors; cannot be checked"], SEE_A5,
     "All (flag test only)"),
    ("DQ13", "Validity", "Average_Daily_Rate",
     ["Average_Daily_Rate of ", ("=", f'IF(MIN({MINADR})<0,"−"&TEXT(ABS(MIN({MINADR})),"0.00"),'
                                      f'TEXT(MIN({MINADR}),"0.00"))'),
      " (", T(f'COUNTIF({MINADR},"<0")'), ") and ", T("ADR_MAX", "#,##0"), " (", T(f"COUNTIF({MINADR},ADR_MAX)"),
      ", about ", T(SV["adr_ratio"], "0"), " times the median of ", T(SV["adr_median"], "0.000"), ")"],
     f'=COUNTIF({MINADR},"<0")+COUNTIF({MINADR},ADR_MAX)', 2, "Flagged invalid (flag_invalid), kept",
     ["A negative price is impossible; ", T("ADR_MAX", "#,##0"), " is almost certainly an entry error"], SEE_A5,
     "All (flag test only)"),
    ("DQ14", "Validity", "Average_Daily_Rate",
     ["Average_Daily_Rate = 0 (", T(SV["zr_comp"]), " of them Complementary)"], "=COUNTIF(rng_flag_zero_rate,1)", 1959,
     "Flagged (flag_zero_rate), kept",
     ["Many are complimentary stays or stays of zero nights; these are not errors"],
     ["Exclude them only for revenue or price analysis"], "None (no price question)"),
    ("DQ15", "Validity", "stays_in_weekend_nights, stays_in_week_nights",
     ["Zero nights (weekend + week nights = 0); ",
      ("=", f'IF({SV["zn_zr"]}=COUNTIF(rng_flag_zero_nights,1),"all have a rate of 0",'
            f'TEXT({SV["zn_zr"]},"#,##0")&" have a rate of 0")')],
     "=COUNTIF(rng_flag_zero_nights,1)", 715, "Flagged (flag_zero_nights), kept",
     ["Possibly day use or same-day changes; cannot be checked"],
     ["Excluding them moves the overall rate from ", T(BASE, "0.0%"), " to ", T(f"Sensitivity!$D${S1['zn']}", "0.0%")],
     "All (very small effect)"),
    ("DQ16", "Validity", "required_car_parking_spaces", ["required_car_parking_spaces = 8"],
     '=COUNTIF(rng_parking,">=8")', 2, "Flagged outlier (flag_outlier)", ["Implausible for one booking"], SEE_A5,
     "All (flag test only)"),
    ("DQ17", "Uniqueness", "all 28 columns except booking_order_id",
     ["Rows identical on all 28 variables but with different booking IDs"], "=SUM(rng_dup)", 32239,
     "Kept, flagged (dup_flag)",
     [T(SV["dup_go_sh"], "0.0%"), " are Groups or Offline TA/TO bookings (against ", T(SV["go_sh"], "0.0%"),
      " of all bookings), so most likely separate rooms in multi-room bookings"],
     ["Removing them would cut the overall rate from ", T(BASE, "0.0%"), " to ",
      T(f"Sensitivity!$D${S1['dup']}", "0.0%"), "; report patterns, and check each key result with dup_flag = 0"],
     "Q1–Q5"),
    ("DQ18", "Consistency", "is_canceled, reservation_status",
     ["is_canceled = Yes covers ", T('COUNTIF(rng_status,"Canceled")'), " cancellations and ",
      T('COUNTIF(rng_status,"No-Show")'), " no-shows; it ",
      ("=", f'IF(SUMPRODUCT((({RAWC}="Yes")*({R("reservation_status")}<>"Check-Out"))+(({RAWC}="No")*'
            f'({R("reservation_status")}="Check-Out")))=N_ALL,"agrees with reservation_status in every row",'
            f'"does NOT agree with reservation_status in every row")')],
     "=COUNTIF(rng_cancel,1)", 44224,
     'Kept: "cancellation" includes no-shows (state this in the report)', ["That is the dataset's own definition"],
     ["Excluding no-shows lowers the overall rate from ", T(BASE, "0.0%"), " to ",
      T(f"Sensitivity!$D${S1['noshow']}", "0.0%")], "Q1–Q5"),
    ("DQ19", "Timeliness", "arrival_date",
     ["Arrivals cover ", T(SV["date_min"], "d mmmm yyyy"), " to ", T(SV["date_max"], "d mmmm yyyy"),
      ", so 2022 and 2024 are part-years (", T("COUNTIF(rng_arrival_year,2022)"), " / ",
      T("COUNTIF(rng_arrival_year,2023)"), " / ", T("COUNTIF(rng_arrival_year,2024)"), " valid-date rows)"],
     "=COUNTIF(rng_date_valid,1)", 119294, "Note only", ["Comparing yearly totals would be misleading"],
     ["Compare rates or months, not yearly booking totals"], "Q1"),
    ("DQ20", "Accuracy, fitness for use (Vidgen, Kirshner & Tan 2019, p. 58)",
     "required_car_parking_spaces, reserved_room_type, assigned_room_type",
     ["Some fields look as if they were filled in at or after check-in. ", T(SV["park_c"]), " of ", T(SV["park_n"]),
      " bookings with a parking space were cancelled, and only ", T(SV["rc_rate"], "0.0%"),
      " of bookings with a changed room type (", T(SV["rc_n"]), " rows). Cancelled guests never arrive, so they never "
      "get a parking space or a room change"],
     F(T(SV["park_n"]), " / ", T(SV["rc_n"])), "7,416 / 14,917", 'Kept; marked "use with caution"',
     ["These fields partly record the outcome rather than cause it; the paper warns that guests often change such "
      "details at check-in (p. 45)"], ["Do NOT present parking or room changes as reasons guests cancel"],
     "Any using parking or room changes"),
]
assert len(DQ) == 20
DQ_ROW = {}
for i, (did, dim, colnames, found, rows_f, exp, done, why, effect, qs) in enumerate(DQ):
    rr = DQ_H + 1 + i
    DQ_ROW[did] = rr
    band = i % 2 == 1

    def txt(parts):
        return F(*parts) if any(not isinstance(p, str) for p in parts) else "".join(parts)
    body_cell(dq, rr, 1, did, None, band, bold=True)
    body_cell(dq, rr, 2, dim, None, band, wrap=True)
    body_cell(dq, rr, 3, colnames, None, band, wrap=True)
    body_cell(dq, rr, 4, txt(found), None, band, wrap=True)
    body_cell(dq, rr, 5, rows_f, FMT_N, band, align=Alignment(horizontal="right", vertical="top"))
    c = body_cell(dq, rr, 6, exp, FMT_N if isinstance(exp, int) else None, band,
                  align=Alignment(horizontal="right", vertical="top"))
    c.font = F_EXP
    body_cell(dq, rr, 7, f'=IF(E{rr}=F{rr},"{OK}","{BAD}")', None, band, align=A_CTR)
    body_cell(dq, rr, 8, done, None, band, wrap=True)
    body_cell(dq, rr, 9, txt(why), None, band, wrap=True)
    body_cell(dq, rr, 10, txt(effect), None, band, wrap=True)
    body_cell(dq, rr, 11, qs, None, band, wrap=True)
    dq.row_dimensions[rr].height = 92
DQ_END = DQ_H + len(DQ)
match_rules(dq, f"G{DQ_H + 1}:G{DQ_END}")
put(dq, f"A{DQ_END + 1}", "Blue = expected value typed from the brief; black = live formula. Paper = Antonio, de Almeida "
                          "& Nunes (2019).", font=F_NOTE)
for col_letter, w in zip("ABCDEFGHIJK", [7, 16, 18, 42, 13, 13, 9, 30, 36, 36, 16]):
    dq.column_dimensions[col_letter].width = w
dq.freeze_panes = f"B{DQ_H + 1}"
presentation(dq, "C00000", area=f"A1:K{DQ_END + 1}")

# ================================================================ Checks
ck_ws = ws_ch
title(ck_ws, "A1", "Checks: live TRUE/FALSE tests")
put(ck_ws, "A2", "The banner in A3 reads ALL CHECKS PASS only when every result below is TRUE (an error value also "
                 "fails it).", font=F_SMALL)
CH_H = 6
header_row(ck_ws, CH_H, 1, ["#", "Check", "Value", "Compared with", "Result"], height=22)
CK = []
dqm = f"Data_Quality_Log!G{DQ_H + 1}:G{DQ_END}"
CK.append(("Data_Quality_Log: every Match? is OK", f'=COUNTIF({dqm},"{OK}")', f"=ROWS({dqm})", "=C{r}=D{r}"))
for key, *_r in BLOCKS:
    info = OV_ROWS[key]
    t = info["total"]
    CK.append((f"OneVar block {key}: total bookings = 119,390 = COUNTA(IDs)", f"=OneVar_Summary!B{t}", 119390,
               "=AND(C{r}=D{r},C{r}=N_ALL)"))
    CK.append((f"OneVar block {key}: total cancelled = 44,224 = SUM(cancel_flag)", f"=OneVar_Summary!C{t}", 44224,
               "=AND(C{r}=D{r},C{r}=C_ALL)"))
mo = OV_ROWS["arrival_month"]
CK.append(("arrival_month: the 12 months hold 119,294 valid-date bookings", f"=SUM(OneVar_Summary!B{mo['first']}:"
           f"B{mo['last'] - 1})", 119294, "=C{r}=D{r}"))
CK.append(("arrival_month: the Invalid date row holds 96 bookings", f"=OneVar_Summary!B{mo['last']}", 96, "=C{r}=D{r}"))
for field, token, exp in (("country", "NULL", 488), ("agent", "NULL", 16340), ("company", "NULL", 112593),
                          ("children", "NA", 4)):
    CK.append((f'{field}: Raw_Data text "{token}" count = Clean_Data blank count = {exp:,}',
               f'=COUNTIF({R(field)},"{token}")', f"=COUNTBLANK(Clean_Data!${COL[field]}$2:${COL[field]}${LAST})",
               f"=AND(C{{r}}=D{{r}},C{{r}}={exp})"))
CK.append(("Clean_Data's first 29 columns differ from Raw_Data in exactly 129,425 cells",
           f"=SUMPRODUCT(--(Raw_Data!$A$2:$AC${LAST}<>Clean_Data!$A$2:$AC${LAST}))", 129425, "=C{r}=D{r}"))
for flag, exp in (("flag_invalid", 186), ("flag_undefined", 5), ("flag_outlier", 241), ("flag_zero_rate", 1959),
                  ("flag_zero_nights", 715), ("flag_any_quality", 428), ("dup_flag", 32239)):
    rng_name = "rng_dup" if flag == "dup_flag" else f"rng_{flag}"
    CK.append((f"Sum of {flag} = {exp:,}", f"=SUM({rng_name})", exp, "=C{r}=D{r}"))
CK.append(("is_canceled agrees with reservation_status in every row (119,390 of 119,390)",
           f'=SUMPRODUCT((({RAWC}="Yes")*({R("reservation_status")}<>"Check-Out"))+(({RAWC}="No")*'
           f'({R("reservation_status")}="Check-Out")))', 119390, "=C{r}=D{r}"))
CK.append(("No #VALUE!, #N/A, #DIV/0! or other error anywhere in tblClean",
           f"=SUMPRODUCT(--ISERROR(Clean_Data!$A$2:${LASTCOL}${LAST}))", 0, "=C{r}=D{r}"))
CK.append(("tblClean has 119,390 data rows", "=ROWS(rng_id)", 119390, "=AND(C{r}=D{r},COUNTA(rng_id)=D{r})"))
CK.append(("Raw_Data has 119,390 data rows", f"=COUNTA({R('booking_order_id')})", 119390, "=C{r}=D{r}"))
CK.append(("Sensitivity Part 1: every overall rate matches the brief", f'=COUNTIF(Sensitivity!G{S1_H + 1}:G{S1_END},"{OK}")',
           f"=ROWS(Sensitivity!G{S1_H + 1}:G{S1_END})", "=C{r}=D{r}"))
rows_ck = []
for i, (lab, left, right, test) in enumerate(CK):
    rr = CH_H + 1 + i
    band = i % 2 == 1
    body_cell(ck_ws, rr, 1, i + 1, "0", band, align=A_CTR)
    body_cell(ck_ws, rr, 2, lab, None, band, wrap=True)
    body_cell(ck_ws, rr, 3, left, FMT_N, band)
    body_cell(ck_ws, rr, 4, right, FMT_N, band)
    body_cell(ck_ws, rr, 5, test.format(r=rr), None, band, align=A_CTR)
CH_END = CH_H + len(CK)
RES = f"$E${CH_H + 1}:$E${CH_END}"
bool_rules(ck_ws, f"E{CH_H + 1}:E{CH_END}")
put(ck_ws, "A3", f'=IF(AND(COUNTIF({RES},FALSE)=0,COUNTIF({RES},TRUE)=ROWS({RES})),"ALL CHECKS PASS","CHECK FAILED")',
    font=Font(name=FONT, size=20, bold=True), align=Alignment(horizontal="left", vertical="center"))
ck_ws.row_dimensions[3].height = 32
ck_ws.conditional_formatting.add("A3", FormulaRule(formula=['A3="ALL CHECKS PASS"'],
                                                   font=Font(name=FONT, size=20, bold=True, color="006100")))
ck_ws.conditional_formatting.add("A3", FormulaRule(formula=['A3<>"ALL CHECKS PASS"'],
                                                   font=Font(name=FONT, size=20, bold=True, color="C00000")))
put(ck_ws, "A4", f'=COUNTIF({RES},TRUE)&" of "&ROWS({RES})&" checks are TRUE."', font=F_SMALL)
s2m = f"Sensitivity!G{S2_FIRST}:G{S2_END}"
put(ck_ws, f"A{CH_END + 2}", F("Information (not in the banner): ", ("=", f'COUNTIF({s2m},"{OK}")'), " of ",
                              ("=", f'COUNTIF({s2m},"{OK}")+COUNTIF({s2m},"{BAD}")'),
                              " brief values for the largest category change match (Sensitivity Part 2). ",
                              ("=", f'IF(COUNTIF({s2m},"{BAD}")>0,"The rest differ from the brief by rounding only; '
                                    f'see the Why column there.","")')), font=F_NOTE)
for col_letter, w in zip("ABCDE", [5, 72, 14, 14, 10]):
    ck_ws.column_dimensions[col_letter].width = w
ck_ws.freeze_panes = f"A{CH_H + 1}"
presentation(ck_ws, "C00000", landscape=False, area=f"A1:E{CH_END + 2}")

# ================================================================ save
wb.active = 0
print(f"Saving {OUT} ...", flush=True)
wb.save(OUT)
print(f"Done in {time.time() - t0:.0f}s; {len(CK)} checks; OneVar rows to {OV_END}; Sensitivity rows to {S4_END}.")
