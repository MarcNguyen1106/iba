# Figure 2 FINAL (v5) reference implementation (tested with openpyxl 3.1.5; rendered with LibreOffice using Carlito = Calibri metrics).
# Design lessons applied: Knaflic (2015) ch. 2 and ch. 8, Duarte (2014), Lecture 6, Berinato (setup / conflict / resolution).
# FINAL fixes: (1) Over 365 band faded (2.6% of bookings, 59% non-refundable) so the eye reads the 0-365 rise first;
# (2) axis labels cut to band + cancelled count (+ non-refundable share) to reduce number density;
# (3) legend names the grey series as a subset with its size: 'No-deposit bookings only (88% of all bookings)'.
# In the real workbook every input comes from Q2_Summary Table D formulas; numbers here are typed only for the render test.
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList, DataLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import DataPoint
from openpyxl.drawing.text import (Paragraph, ParagraphProperties, CharacterProperties, Font,
                                   RichTextProperties, RegularTextRun)

FOCUS = '1F5AA6'      # strong blue: the answer (all bookings)
CONTEXT = 'A6A6A6'    # mid grey: the check (no-deposit bookings)
CONTEXT_TXT = '7F7F7F'
LIGHT_FOCUS = 'BDD0EA'  # pale blue: de-emphasised bars (Over 365 band; presentation 'where' view)
LIGHT_FOCUS_TXT = '5B7DB8'  # readable muted blue for labels on de-emphasised blue bars
LIGHT_CONTEXT = 'DCDCDC'  # pale grey: de-emphasised no-deposit bar (Over 365 band)
HEAD = '262626'
SUB = '595959'
AXIS_TXT = '595959'
AXIS_LINE = 'BFBFBF'


def cp(sz, b=False, color=AXIS_TXT):
    return CharacterProperties(sz=sz, b=b, solidFill=color, latin=Font(typeface='Calibri'))


def rich(sz, b=False, color=AXIS_TXT, rot=None):
    bp = RichTextProperties(rot=rot) if rot is not None else RichTextProperties()
    return RichText(bodyPr=bp, p=[Paragraph(pPr=ParagraphProperties(defRPr=cp(sz, b, color)), endParaRPr=cp(sz, b, color))])


def run(text, sz, b, color):
    return RegularTextRun(rPr=cp(sz, b, color), t=text)


def story_title(headline):
    """One-line action headline, single format (renders identically in Excel and LibreOffice), placed top-left."""
    p1 = Paragraph(pPr=ParagraphProperties(algn='l', defRPr=cp(1200, True, HEAD)), r=[run(headline, 1200, True, HEAD)])
    t = Title(tx=Text(rich=RichText(bodyPr=RichTextProperties(), p=[p1])), overlay=False)
    t.layout = Layout(manualLayout=ManualLayout(xMode='edge', yMode='edge', x=0.01, y=0.01))
    return t


def axis_title(text):
    """Axis title pushed to the background: regular weight, grey (Knaflic 2015, ch. 8)."""
    p1 = Paragraph(pPr=ParagraphProperties(defRPr=cp(900, False, AXIS_TXT)), r=[run(text, 900, False, AXIS_TXT)])
    return Title(tx=Text(rich=RichText(bodyPr=RichTextProperties(), p=[p1])), overlay=False)


REPORT_HEADLINE = 'The further ahead tourists book, the more likely they are to cancel'


def _style_series(s, fill, label_color, bold):
    s.graphicalProperties.solidFill = fill
    s.graphicalProperties.line.noFill = True
    s.dLbls = DataLabelList()
    s.dLbls.showVal = True
    for k in ('showSerName', 'showCatName', 'showLegendKey', 'showPercent'):
        setattr(s.dLbls, k, False)
    s.dLbls.position = 'outEnd'
    s.dLbls.numFmt = '0.0%'
    s.dLbls.txPr = rich(850, bold, label_color)


def _fade_point(series, idx, fill, label_color):
    """Push one bar into the background: pale fill and a lighter, regular-weight label (Knaflic 2015, p. 195)."""
    pt = DataPoint(idx=idx)
    pt.graphicalProperties.solidFill = fill
    pt.graphicalProperties.line.noFill = True
    series.dPt.append(pt)
    lbl = DataLabel(idx=idx, showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False)
    lbl.position = 'outEnd'
    lbl.numFmt = '0.0%'
    lbl.txPr = rich(850, False, label_color)
    series.dLbls.dLbl.append(lbl)


def build_chart(data_ws, header_row, first_row, last_row, cat_col, all_col, nd_col,
                headline=REPORT_HEADLINE, include_nd=True, highlight_points=None, legend=True, fade_points=(5,)):
    ch = BarChart()
    ch.type = 'col'
    ch.grouping = 'clustered'
    max_col = nd_col if include_nd else all_col
    ch.add_data(Reference(data_ws, min_col=all_col, max_col=max_col, min_row=header_row, max_row=last_row), titles_from_data=True)
    ch.set_categories(Reference(data_ws, min_col=cat_col, min_row=first_row, max_row=last_row))
    ch.gapWidth = 40 if include_nd else 60
    ch.overlap = -12 if include_nd else 0
    _style_series(ch.series[0], FOCUS, FOCUS, True)
    if include_nd:
        _style_series(ch.series[1], CONTEXT, CONTEXT_TXT, False)
    if highlight_points is None:   # report and setup/conflict views: fade the small, deposit-driven Over 365 band
        for idx in fade_points:
            _fade_point(ch.series[0], idx, LIGHT_FOCUS, LIGHT_FOCUS_TXT)
            if include_nd:
                _fade_point(ch.series[1], idx, LIGHT_CONTEXT, CONTEXT_TXT)
    if highlight_points is not None:   # presentation 'resolution' view: keep highlighted bars blue, fade the rest
        ch.series[0].graphicalProperties.solidFill = LIGHT_FOCUS
        for idx in highlight_points:
            pt = DataPoint(idx=idx)
            pt.graphicalProperties.solidFill = FOCUS
            pt.graphicalProperties.line.noFill = True
            ch.series[0].dPt.append(pt)
    ch.y_axis.scaling.min = 0
    ch.y_axis.scaling.max = 0.8
    ch.y_axis.majorUnit = 0.2
    ch.y_axis.number_format = '0%'
    ch.y_axis.majorGridlines = None
    ch.y_axis.delete = True
    ch.x_axis.title = axis_title('Lead time (days between booking and arrival)')
    ch.x_axis.txPr = rich(800, False, AXIS_TXT, rot=0)
    ch.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=AXIS_LINE, w=9525))
    ch.x_axis.delete = False
    ch.x_axis.tickLblSkip = 1
    ch.x_axis.tickMarkSkip = 1
    if legend:
        ch.legend.position = 't'
        ch.legend.txPr = rich(900, False, '404040')
    else:
        ch.legend = None
    ch.title = story_title(headline)
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.width = 15.9
    ch.height = 10.0
    return ch


if __name__ == '__main__':
    rows = [('0–7 days', 1902, 0.0963233, 0.0916459), ('8–30 days', 5283, 0.2786392, 0.2604404), ('31–90 days', 11141, 0.3769837, 0.3218984),
            ('91–180 days', 11821, 0.4471047, 0.3486074), ('181–365 days', 11947, 0.5545395, 0.3982395), ('Over 365 days', 2130, 0.6766201, 0.20625)]
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Data'
    ws.append(['Category', 'All bookings', 'No-deposit bookings only (88% of all bookings)'])
    for b, c, a, d in rows:
        ws.append([f'{b}\n{c:,} cancelled' + ('\n59% non-refundable' if b.startswith('Over') else ''), a, d])
    for r in ws.iter_rows(min_row=2, max_row=7, min_col=2, max_col=3):
        for c in r:
            c.number_format = '0.0%'
    f = wb.create_sheet('Figure2'); f.add_chart(build_chart(ws, 1, 2, 7, 1, 2, 3), 'A1')
    for s_ in (f,):
        s_.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save('figure2_FINAL_reference.xlsx')
