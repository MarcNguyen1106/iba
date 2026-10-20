#!/usr/bin/env python3
"""Draw the redesigned Figure 1 and Figure C1 for the MKTG3007 Trung Nguyen report.

Each figure is drawn on a canvas at its exact printed size, with one axes that
covers the whole canvas in millimetres (origin top-left, y pointing down).
Marker diameters, bar heights and gaps are real millimetres, and line weights
and font sizes are real points, so nothing has to be rescaled in Word.

Usage:  python make_figures.py [out_dir]          (default: ./output)
"""
from pathlib import Path
import sys

import numpy as np
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.patches import Circle, FancyBboxPatch, Polygon, Rectangle  # noqa: E402
from PIL import Image  # noqa: E402

PT = 25.4 / 72  # millimetres per point
DPI = 400

# Palette (brief, section 3). Colour follows the entity in both figures.
TN = "#C0561E"         # Trung Nguyên
COMP = "#3A6EA5"       # competitors
DIST = "#2A9D8F"       # distributor
TN_TINT = "#F6E3D7"
TN_BAND = "#FAEEE7"     # TN_TINT at 60 % on white: a highlight band, not a block
DIST_TINT = "#DDF0EC"
INK = "#1E2329"
TEXT2 = "#5B6470"
MUTED = "#8A939E"
RULE = "#E4E7EB"
BG = "#FFFFFF"
STAGE_RAMP = ["#F6E3D7", "#EBB48F", "#D9814F", "#A9461A"]

DASH_ON, DASH_OFF = 5.0, 3.0  # points
MARKER_R = 3.2 / 2            # marker radius, mm
RING = 1.5 * PT               # white ring around every marker, mm
BAR_RADIUS = 0.8              # corner radius of every bar, mm (both ends)

plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": ["Arial", "Liberation Sans"],
    "lines.scale_dashes": False,  # dash lengths in points, whatever the line width
    "svg.fonttype": "path",
    "svg.hashsalt": "mktg3007",
})


def cap_height(size):
    """Cap height of Arial / Liberation Sans in mm for a size in points."""
    return 0.716 * size * PT


def line_pitch(size, spacing=1.18):
    return size * spacing * PT


class Canvas:
    """A w x h mm figure with a single mm-based axes (y points down)."""

    def __init__(self, w, h):
        self.w, self.h = w, h
        self.fig = plt.figure(figsize=(w / 25.4, h / 25.4), dpi=DPI)
        self.fig.patch.set_facecolor(BG)
        self.ax = self.fig.add_axes((0, 0, 1, 1))
        self.ax.set_axis_off()
        self.set_height(h)
        self.renderer = self.fig.canvas.get_renderer()
        self.texts = []      # text artists, checked for collisions
        self.marks = []      # (x0, y0, x1, y1, name) boxes of data marks
        self.curves = []     # (xs, ys, name) polylines that text must not cross

    def set_height(self, h):
        self.h = h
        self.fig.set_size_inches(self.w / 25.4, h / 25.4)
        self.ax.set_xlim(0, self.w)
        self.ax.set_ylim(h, 0)

    # ---- text -------------------------------------------------------------
    def text(self, x, y, s, size, color=INK, weight="normal", ha="left",
             va="baseline", z=6, check=True, **kw):
        t = self.ax.text(x, y, s, fontsize=size, color=color, fontweight=weight,
                         ha=ha, va=va, zorder=z, **kw)
        if check:
            self.texts.append(t)
        return t

    def lines(self, x, first_baseline, rows, size, pitch=None, **kw):
        """Stack single-line texts on an exact baseline grid."""
        pitch = pitch or line_pitch(size)
        return [self.text(x, first_baseline + i * pitch, r, size, **kw)
                for i, r in enumerate(rows)]

    def width(self, s, size, weight="normal"):
        t = self.ax.text(0, 0, s, fontsize=size, fontweight=weight)
        w = t.get_window_extent(self.renderer).width / DPI * 25.4
        t.remove()
        return w

    def bbox(self, artist):
        """Artist extent as (x0, y0, x1, y1) in mm, y down."""
        b = artist.get_window_extent(self.renderer)
        k = 25.4 / DPI
        return (b.x0 * k, self.h - b.y1 * k, b.x1 * k, self.h - b.y0 * k)

    # ---- marks ------------------------------------------------------------
    def dot(self, x, y, color, z=4, ring_color=BG):
        self.ax.add_patch(Circle((x, y), MARKER_R + RING, facecolor=ring_color,
                                 edgecolor="none", zorder=z))
        self.ax.add_patch(Circle((x, y), MARKER_R, facecolor=color,
                                 edgecolor="none", zorder=z + 0.01))
        self.marks.append((x - MARKER_R, y - MARKER_R, x + MARKER_R, y + MARKER_R, "dot"))

    def open_ring(self, x, y, color, lw=1.3, n=6, z=4, ring_color=BG):
        """Hollow marker with a dashed outline; its outer edge matches a filled dot."""
        r = MARKER_R - lw * PT / 2
        period = 2 * np.pi * r / PT / n            # n whole dashes round the circle
        on, off = period * DASH_ON / (DASH_ON + DASH_OFF), period * DASH_OFF / (DASH_ON + DASH_OFF)
        self.ax.add_patch(Circle((x, y), MARKER_R + RING, facecolor=ring_color,
                                 edgecolor="none", zorder=z))
        ring = Circle((x, y), r, facecolor=ring_color, edgecolor=color, linewidth=lw,
                      linestyle=(on / 2, (on, off)), zorder=z + 0.01)
        ring.set_capstyle("butt")
        self.ax.add_patch(ring)
        self.marks.append((x - MARKER_R, y - MARKER_R, x + MARKER_R, y + MARKER_R, "ring"))

    def arrow(self, x_tail, x_tip, y, color, lw, head_len, head_w, dashed=False, z=4.005):
        """Horizontal arrow: straight shaft of lw points and a filled triangular
        head whose tip sits exactly on x_tip. The default z puts it above the
        markers' white rings but below their discs, so the tip meets the disc."""
        d = np.sign(x_tip - x_tail)
        x_base = x_tip - d * head_len
        self.ax.add_patch(Polygon([(x_tip, y), (x_base, y - head_w / 2), (x_base, y + head_w / 2)],
                                  closed=True, facecolor=color, edgecolor="none",
                                  zorder=z + 0.001))
        tuck = 0.2  # shaft runs 0.2 mm under the head so no seam shows
        x_join = x_base + d * tuck
        if dashed:
            # Whole dashes only: a full dash meets the head and another ends at
            # the tail; the gaps stretch slightly to make the length fit.
            length = abs(x_base - x_tail) / PT
            n = max(2, round((length + DASH_OFF) / (DASH_ON + DASH_OFF)))
            off = (length - n * DASH_ON) / (n - 1)
            seq = [DASH_ON + tuck / PT, off] + [DASH_ON, off] * (n - 2) + [DASH_ON, off]
            self.ax.plot([x_join, x_tail], [y, y], color=color, lw=lw,
                         linestyle=(0, seq), dash_capstyle="butt", zorder=z)
        else:
            self.ax.plot([x_tail, x_join], [y, y], color=color, lw=lw,
                         solid_capstyle="butt", zorder=z)
        self.marks.append((min(x_tail, x_tip), y - head_w / 2, max(x_tail, x_tip), y + head_w / 2, "arrow"))

    def bar(self, x0, x1, yc, h, color, z=3, name="bar", **kw):
        r = min(BAR_RADIUS, h / 2)
        self.ax.add_patch(FancyBboxPatch((x0, yc - h / 2), x1 - x0, h,
                                         boxstyle=f"round,pad=0,rounding_size={r}",
                                         facecolor=color, zorder=z, **{"edgecolor": "none", **kw}))
        self.marks.append((x0, yc - h / 2, x1, yc + h / 2, name))

    def diamond(self, x, y, half, color=INK, z=5):
        self.ax.add_patch(Polygon([(x, y - half), (x + half, y), (x, y + half), (x - half, y)],
                                  closed=True, facecolor=color, edgecolor="none", zorder=z))
        self.marks.append((x - half, y - half, x + half, y + half, "diamond"))

    def hline(self, x0, x1, y, color=RULE, lw=0.7, z=1, **kw):
        self.ax.plot([x0, x1], [y, y], color=color, lw=lw, solid_capstyle="butt", zorder=z, **kw)

    def vguide(self, x, y0, y1, color=RULE, lw=0.7, z=1, gap=0.9, avoid=None):
        """Vertical guide line that breaks around any text it would cross."""
        blocked = []
        for t in (self.texts if avoid is None else avoid):
            bx0, by0, bx1, by1 = self.bbox(t)
            if bx0 - gap <= x <= bx1 + gap:
                blocked.append((by0 - gap, by1 + gap))
        blocked.sort()
        segs, cur = [], y0
        for b0, b1 in blocked:
            if b1 <= cur or b0 >= y1:
                continue
            if b0 > cur:
                segs.append((cur, b0))
            cur = max(cur, b1)
        if cur < y1:
            segs.append((cur, y1))
        for s0, s1 in segs:
            if s1 - s0 > 1.5:   # no slivers between neighbouring labels
                self.ax.plot([x, x], [s0, s1], color=color, lw=lw,
                             solid_capstyle="butt", zorder=z)

    # ---- checks and output ------------------------------------------------
    def check_layout(self, name):
        """Report text that overlaps other text, a mark, a curve or the canvas edge."""
        problems = []
        boxes = [(self.bbox(t), t.get_text()) for t in self.texts]
        for i, (a, sa) in enumerate(boxes):
            if a[0] < 0 or a[1] < 0 or a[2] > self.w or a[3] > self.h:
                problems.append(f"outside canvas: {sa!r}")
            for b, sb in boxes[i + 1:]:
                if a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]:
                    problems.append(f"text/text: {sa!r} x {sb!r}")
            for m in self.marks:
                if a[0] < m[2] and m[0] < a[2] and a[1] < m[3] and m[1] < a[3]:
                    if not getattr(self, "inside_ok", set()) or sa not in self.inside_ok:
                        problems.append(f"text/{m[4]}: {sa!r}")
            for xs, ys, cname in self.curves:
                inside = (xs >= a[0] - 0.6) & (xs <= a[2] + 0.6) & (ys >= a[1] - 0.6) & (ys <= a[3] + 0.6)
                if inside.any():
                    problems.append(f"text/{cname}: {sa!r}")
        sizes = sorted({t.get_fontsize() for t in self.texts})
        print(f"[{name}] {self.w:.1f} x {self.h:.1f} mm, font sizes {sizes} pt")
        print(f"[{name}] layout check: {'OK' if not problems else ''}")
        for p in problems:
            print("   -", p)
        return problems

    def save(self, out_dir, stem):
        png, svg = out_dir / f"{stem}.png", out_dir / f"{stem}.svg"
        self.fig.savefig(png, dpi=DPI, facecolor=BG)
        self.fig.savefig(svg, facecolor=BG)
        # Opaque RGB keeps the PNG small; the pHYs chunk records the print size.
        Image.open(png).convert("RGB").save(png, dpi=(DPI, DPI), optimize=True)
        plt.close(self.fig)
        return png, svg


# ===========================================================================
# Figure 1: entry routes in Australian coffee
# ===========================================================================
def figure1(out_dir, band=TN_BAND, header_color=TEXT2, stem="figure1"):
    W = 142.0
    c = Canvas(W, 130.0)

    NAME, QUAL, NOTE, HEAD, SMALL = 9.0, 7.5, 7.5, 8.5, 7.5
    note_pitch = line_pitch(NOTE, 1.16)
    head_pitch = line_pitch(HEAD, 1.14)

    xs = np.linspace(36.5, 122.5, 5)      # five evenly spaced entry modes
    EXPORT, DISTRIB, LICENCE, JV, OWN = xs

    # --- direction of travel ------------------------------------------------
    y_dir = 3.2
    x_start = xs[0] - 10.5
    t = c.text(x_start, y_dir, "Rising resource commitment, risk and control", SMALL, color=MUTED)
    x_line = c.bbox(t)[2] + 2.0
    y_mid = y_dir - 0.36 * SMALL * PT * 1.0
    c.ax.plot([x_line, xs[-1] + 9.0], [y_mid, y_mid], color=MUTED, lw=0.8,
              solid_capstyle="butt", zorder=2)
    c.ax.add_patch(Polygon([(xs[-1] + 10.5, y_mid), (xs[-1] + 8.9, y_mid - 0.75),
                            (xs[-1] + 8.9, y_mid + 0.75)], closed=True,
                           facecolor=MUTED, edgecolor="none", zorder=2))

    # --- column headers, bottom-aligned on one baseline ----------------------
    headers = [["Export via", "retailer or", "importer"], ["Exclusive", "distributor"],
               ["Licensing or", "franchising"], ["Joint", "venture"], ["Full", "ownership"]]
    hdr_last = y_dir + 5.7 + 2 * head_pitch
    for x, rows in zip(xs, headers):
        c.lines(x, hdr_last - (len(rows) - 1) * head_pitch, rows, HEAD,
                pitch=head_pitch, weight="bold", ha="center", color=header_color)

    # --- rows ------------------------------------------------------------------
    P = 11.6                      # row pitch
    note_gap = MARKER_R + 1.25    # dot edge to top of note caps

    def note(x, y, rows, color):
        c.lines(x, y + note_gap + cap_height(NOTE), rows, NOTE, pitch=note_pitch,
                color=color, ha="center")

    def row_label(y, name, qual, focus):
        c.text(0, y + cap_height(NAME) / 2, name, NAME,
               color=INK if focus else TEXT2, weight="bold" if focus else "normal")
        c.text(0, y + note_gap + cap_height(NOTE), qual, QUAL, color=INK if focus else TEXT2)

    c.text(0, hdr_last, "Competitors", SMALL, color=TEXT2, weight="bold")
    y0 = hdr_last + 6.1
    comp_rows = [y0 + i * P for i in range(5)]
    nestle, jde, lavazza, starbucks, king = comp_rows

    row_label(nestle, "Nestlé", "Switzerland", False)
    c.dot(OWN, nestle, COMP)
    note(OWN, nestle, ["Owns Gympie plant: over half", "of coffee drunk in Australia"], TEXT2)

    row_label(jde, "JDE Peet's", "Netherlands", False)
    c.dot(OWN, jde, COMP)
    note(OWN, jde, ["Bought roaster", "Campos (2021)"], TEXT2)

    HL, HW = 2.0, 1.9            # competitor arrowhead length, width (mm)
    row_label(lavazza, "Lavazza", "Italy", False)
    c.dot(DISTRIB, lavazza, COMP)
    c.dot(OWN, lavazza, COMP)
    c.arrow(DISTRIB + MARKER_R + RING, OWN - MARKER_R, lavazza, COMP, 1.6, HL, HW)
    note(DISTRIB, lavazza, ["Blue Pod, exclusive", "for 10+ years"], TEXT2)
    note(OWN, lavazza, ["Bought Blue Pod", "(2018)"], TEXT2)

    row_label(starbucks, "Starbucks", "US", False)
    c.dot(OWN, starbucks, COMP)
    c.dot(LICENCE, starbucks, COMP)
    c.arrow(OWN - MARKER_R - RING, LICENCE + MARKER_R, starbucks, COMP, 1.6, HL, HW, dashed=True)
    note(OWN, starbucks, ["Owned stores 2000–08;", "61 of 84 closed"], TEXT2)
    note(LICENCE, starbucks, ["Licensee Withers", "Group (2014)"], TEXT2)

    row_label(king, "King Coffee", "Vietnam", False)
    c.dot(EXPORT, king, COMP)
    note(EXPORT, king, ["Through Costco", "(2024)"], TEXT2)

    # --- Trung Nguyên block ----------------------------------------------------
    sep_y = king + note_gap + cap_height(NOTE) + note_pitch + 2.9
    tn_label_y = sep_y + 4.2
    tn_now = tn_label_y + 5.5
    tn_prop = tn_now + P
    c.text(0, tn_label_y, "Trung Nguyên", SMALL, color=TEXT2, weight="bold")

    row_label(tn_now, "Trung Nguyên", "now", True)
    c.dot(EXPORT, tn_now, TN, ring_color=band or BG)
    c.open_ring(LICENCE, tn_now, TN, ring_color=band or BG)
    note(EXPORT, tn_now, ["G7, Legend instant", "via importers"], INK)
    note(LICENCE, tn_now, ["Melbourne café:", "announced, unopened"], INK)

    row_label(tn_prop, "Trung Nguyên", "proposed", True)
    c.dot(DISTRIB, tn_prop, TN, ring_color=band or BG)
    c.dot(OWN, tn_prop, TN, ring_color=band or BG)
    c.arrow(DISTRIB + MARKER_R + RING, OWN - MARKER_R, tn_prop, TN, 2.4, 2.7, 2.7)
    note(DISTRIB, tn_prop, ["Stage 1 (2027):", "distributor +", "Amazon store"], INK)
    note((DISTRIB + OWN) / 2, tn_prop, ["Stage 2 (Sep 2027):", "Woolworths and Coles"], INK)
    note(OWN, tn_prop, ["Stage 3 (2029,", "if triggers met):", "sales subsidiary"], INK)

    tn_bottom = tn_prop + note_gap + cap_height(NOTE) + 2 * note_pitch + 0.9
    if band:
        band_top, band_bottom = sep_y, tn_bottom + 2.2
        c.ax.add_patch(Rectangle((0, band_top), W, band_bottom - band_top,
                                 facecolor=band, edgecolor="none", zorder=0))
    else:
        band_bottom = tn_bottom + 2.2
        c.hline(0, W, sep_y, color=RULE, lw=0.7)

    # --- legend: one row, centred ---------------------------------------------
    leg_y = band_bottom + 4.2
    leg_base = leg_y + cap_height(SMALL) / 2
    items = [("dot", COMP, "Competitor"), ("dot", TN, "Trung Nguyên"),
             ("ring", TN, "Announced, not opened"), ("solid", TEXT2, "Later move"),
             ("dashed", TEXT2, "Retreat")]
    glyph_w = {"dot": 2 * MARKER_R, "ring": 2 * MARKER_R, "solid": 8.0, "dashed": 8.0}
    gap_glyph, gap_item = 1.6, 5.0
    total = sum(glyph_w[k] + gap_glyph + c.width(s, SMALL) for k, _, s in items) + gap_item * (len(items) - 1)
    x = (W - total) / 2
    for kind, col, label in items:
        g = glyph_w[kind]
        if kind == "dot":
            c.dot(x + MARKER_R, leg_y, col)
        elif kind == "ring":
            c.open_ring(x + MARKER_R, leg_y, col)
        elif kind == "solid":
            c.arrow(x, x + g, leg_y, col, 1.6, HL, HW)
        else:
            c.arrow(x + g, x, leg_y, col, 1.6, HL, HW, dashed=True)
        x += g + gap_glyph
        c.text(x, leg_base, label, SMALL, color=TEXT2)
        x += c.width(label, SMALL) + gap_item

    foot_y = leg_y + 5.5
    c.text(0, foot_y, "Mode classification is the author's, based on the sources in Appendix B.",
           7.0, color=MUTED)
    c.set_height(foot_y + 1.3)

    # --- column guides, behind everything and broken around notes -------------
    guide_top = hdr_last + 2.2
    guide_bottom = tn_prop
    for x in xs:
        c.vguide(x, guide_top, guide_bottom)

    c.check_layout("figure1")
    return c.save(out_dir, stem)


# ===========================================================================
# Figure C1: staged roadmap + international distribution life cycle
# ===========================================================================
def distributor_curve(t):
    return 0.30 + 0.55 * np.exp(-((t - 12) / 14) ** 2)


def owner_curve(t, mid):
    return 0.08 + 0.85 / (1 + np.exp(-(t - mid) / 5.5))


def owner_midpoint_for_crossover(t_cross):
    """Midpoint of the logistic curve that makes the two curves cross at t_cross."""
    d = distributor_curve(t_cross)
    return t_cross + 5.5 * np.log(0.85 / (d - 0.08) - 1)


def crossing(mid):
    """Numerical intersection of the two curves after the distributor's peak."""
    f = lambda t: distributor_curve(t) - owner_curve(t, mid)  # noqa: E731
    lo, hi = 12.0, 47.0
    for _ in range(80):
        m = (lo + hi) / 2
        lo, hi = (m, hi) if f(m) > 0 else (lo, m)
    return (lo + hi) / 2


def figure_c1(out_dir, year_bands=False, teal_fill=True):
    W = 159.0
    c = Canvas(W, 125.0)
    X0, X1 = 23.0, 156.5          # Jan 2027 .. end of Dec 2030
    MONTHS = 48

    def tx(t):
        return X0 + (X1 - X0) * t / MONTHS

    LABEL, ANNO, SMALL = 8.5, 8.0, 7.5
    anno_pitch = line_pitch(ANNO, 1.15)

    # Gates sit on stage boundaries: end of the named month (t in months from Jan 2027).
    T_GATE0, T_GATE1, T_GATE2, T_TRIGGER = 3, 9, 15, 27
    mid = owner_midpoint_for_crossover(T_TRIGGER)
    t_cross = crossing(mid)
    print(f"[figureC1] owner-curve midpoint {mid:.2f} months -> crossover at t = {t_cross:.3f} "
          f"(trigger check at t = {T_TRIGGER})")

    # --- milestones -------------------------------------------------------------
    y_name, y_date, y_diamond = 2.9, 6.1, 9.9
    gates = [(T_GATE0, "Gate 0", "Mar 2027"), (T_GATE1, "Gate 1", "Sep 2027"),
             (T_GATE2, "Gate 2", "Mar 2028"), (T_TRIGGER, "Trigger check", "Mar 2029")]
    for t, name, date in gates:
        c.text(tx(t), y_name, name, SMALL, color=INK, ha="center")
        c.text(tx(t), y_date, date, SMALL, color=TEXT2, ha="center")
        c.diamond(tx(t), y_diamond, 1.35)

    # --- roadmap rows -------------------------------------------------------------
    BAR_H, CONTRACT_H, INSET = 5.0, 2.0, 0.35
    rows_y = [17.6, 27.4, 37.2, 47.0]
    contract_y = 56.8           # same pitch as the stage rows

    def row_label(y, rows):
        n = len(rows)
        first = y + cap_height(LABEL) / 2 - (n - 1) * line_pitch(LABEL, 1.12) / 2
        c.lines(0, first, rows, LABEL, pitch=line_pitch(LABEL, 1.12), color=INK)

    def outside(x, y, rows):
        n = len(rows)
        first = y + cap_height(ANNO) / 2 - (n - 1) * anno_pitch / 2
        c.lines(x, first, rows, ANNO, pitch=anno_pitch, color=INK)

    stages = [
        ("Preparation", 0, 3, STAGE_RAMP[0], ["IP Australia filing, distributor", "chosen, labels approved"]),
        ("Stage 1", 3, 9, STAGE_RAMP[1], ["Distributor (NSW, VIC):", "Asian grocers, IGA", "+ Amazon.com.au"]),
        ("Stage 2", 9, 27, STAGE_RAMP[2], ["Woolworths and Coles listings", "via the distributor"]),
        ("Stage 3", 27, 48, STAGE_RAMP[3], None),
    ]
    c.inside_ok = set()
    for (name, t0, t1, colour, desc), y in zip(stages, rows_y):
        row_label(y, [name])
        x0 = tx(t0) + INSET
        if desc is None:
            # Open-ended: bar stops short of the axis end, a chevron carries it on.
            x1 = tx(t1) - 2.9
            c.bar(x0, x1, y, BAR_H, colour, name=name)
            cx0 = x1 + 0.7
            chev = [(cx0, y - BAR_H / 2), (cx0 + 1.1, y - BAR_H / 2), (cx0 + 2.2, y),
                    (cx0 + 1.1, y + BAR_H / 2), (cx0, y + BAR_H / 2), (cx0 + 1.1, y)]
            c.ax.add_patch(Polygon(chev, closed=True, facecolor=colour, edgecolor="none", zorder=3))
            label = "Sales subsidiary; flagship café later"
            pad = 2.2
            assert c.width(label, ANNO) + 2 * pad <= x1 - x0, "Stage 3 label does not fit inside its bar"
            c.text(x0 + pad, y + cap_height(ANNO) / 2, label, ANNO, color=BG)
            c.inside_ok.add(label)
        else:
            x1 = tx(t1) - INSET
            c.bar(x0, x1, y, BAR_H, colour, name=name)
            outside(tx(t1) + 2.0, y, desc)

    # Distributor contract: Apr 2027 - Mar 2029, first six months a trial.
    row_label(contract_y, ["Distributor", "contract"])
    c.bar(tx(3) + INSET, tx(9) - INSET, contract_y, CONTRACT_H, DIST_TINT, name="trial",
          edgecolor=DIST, linewidth=0.6)
    c.bar(tx(9) + INSET, tx(27) - INSET, contract_y, CONTRACT_H, DIST, name="contract")
    outside(tx(27) + 2.0, contract_y, ["Two years; six-month trial;", "buy-out option at expiry"])
    top_panel_bottom = contract_y + 4.2

    # --- lower panel: international distribution life cycle ----------------------
    y_top, y_base = top_panel_bottom + 6.0, top_panel_bottom + 49.0
    VMAX = 1.17

    def vy(v):
        return y_base - (y_base - y_top) * v / VMAX

    t = np.linspace(0, MONTHS, 600)
    d_y = vy(distributor_curve(t))
    o_y = vy(owner_curve(t, mid))
    xs = tx(t)
    if teal_fill:
        c.ax.fill_between(xs, d_y, y_base, facecolor=DIST_TINT, edgecolor="none", alpha=0.6, zorder=1.5)
    c.ax.plot(xs, d_y, color=DIST, lw=2.2, linestyle=(0, (DASH_ON, DASH_OFF)),
              dash_capstyle="butt", solid_joinstyle="round", zorder=3)
    c.ax.plot(xs, o_y, color=TN, lw=2.4, solid_capstyle="butt", solid_joinstyle="round", zorder=3.1)
    c.curves += [(xs, d_y, "distributor curve"), (xs, o_y, "owner curve")]

    # baseline and year ticks
    c.hline(X0, X1, y_base, color=MUTED, lw=0.6, z=2)
    for year, tt in zip((2027, 2028, 2029, 2030), (0, 12, 24, 36)):
        c.ax.plot([tx(tt), tx(tt)], [y_base, y_base + 1.2], color=MUTED, lw=0.6, zorder=2)
        c.text(tx(tt), y_base + 1.6 + cap_height(SMALL), f"Jan {year}", SMALL, color=TEXT2, ha="center")

    c.text(5.0, (y_top + y_base) / 2, "Effort and commitment\n(illustrative)", ANNO, color=TEXT2,
           rotation=90, ha="center", va="center", multialignment="center", linespacing=1.15)

    # Direct labels with a short sample of each curve's line style.
    def curve_label(x_text, y_first, rows, style):
        sx1 = x_text - 1.8
        sx0 = sx1 - 6.0
        sy = y_first - cap_height(ANNO) / 2 + (len(rows) - 1) * anno_pitch / 2
        c.ax.plot([sx0, sx1], [sy, sy], zorder=3, solid_capstyle="butt", dash_capstyle="butt", **style)
        c.lines(x_text, y_first, rows, ANNO, pitch=anno_pitch, color=INK)
        return sx0

    teal_rows = ["Distributor's willingness and ability", "to build Trung Nguyên's share"]
    teal_style = dict(color=DIST, lw=2.2, linestyle=(0, (DASH_ON, DASH_OFF)))
    orange_rows = ["Trung Nguyên's own ability and", "willingness to build share"]
    orange_style = dict(color=TN, lw=2.4)
    block_h = cap_height(ANNO) + anno_pitch

    # distributor label: above its peak, left
    x_text = X0 + 9.0
    w = max(c.width(r, ANNO) for r in teal_rows)
    span = (t >= 0) & (xs <= x_text + w)
    y_last = d_y[span].min() - 3.0 - 0.7          # 3 mm clear of the curve (incl. descenders)
    curve_label(x_text, y_last - anno_pitch, teal_rows, teal_style)

    # Trung Nguyên label: above the right end of its curve, right-aligned to the axis
    w = max(c.width(r, ANNO) for r in orange_rows)
    x_text = X1 - w
    span = xs >= x_text - 8.0
    y_last = o_y[span].min() - 3.0 - 0.7
    curve_label(x_text, y_last - anno_pitch, orange_rows, orange_style)

    # Trigger-check line: ink at 50 %, from its diamond through both panels.
    xt = tx(T_TRIGGER)
    c.ax.plot([xt, xt], [y_diamond + 1.35, y_base], color=INK, lw=0.9, alpha=0.5,
              solid_capstyle="butt", zorder=2)

    # Crossover marker and label (upper right, in the wedge between the curves).
    v_cross = distributor_curve(t_cross)
    c.dot(tx(t_cross), vy(v_cross), INK, z=5)
    label = "Crossover: buy out or open own office"
    lw_ = c.width(label, ANNO)
    lh = cap_height(ANNO) + 0.75     # caps plus descender
    margin = 1.3
    for t_left in np.arange(t_cross + 1.0, 46.0, 0.05):
        top_limit = vy(owner_curve(t_left, mid)) + margin          # just below orange
        bottom_limit = vy(distributor_curve(t_left)) - margin      # just above teal
        if bottom_limit - top_limit >= lh and tx(t_left) + lw_ <= X1:
            break
    base = (top_limit + bottom_limit) / 2 + lh / 2 - 0.75
    c.text(tx(t_left), base, label, ANNO, color=INK)

    # --- footnote -------------------------------------------------------------------
    foot = ["Lower panel adapted from the international distribution life cycle (Arnold and Quelch, in",
            "Czinkota, Ronkainen & Cui 2023, p. 441); curves are illustrative, not data."]
    f0 = y_base + 10.4
    c.lines(0, f0, foot, 7.0, pitch=line_pitch(7.0, 1.2), color=MUTED)
    c.set_height(f0 + line_pitch(7.0, 1.2) + 1.6)

    # Year bands (optional), stage guides broken around labels.
    if year_bands:
        for t0 in (12, 36):
            c.ax.add_patch(Rectangle((tx(t0), y_diamond + 3.0), tx(t0 + 12) - tx(t0),
                                     y_base - y_diamond - 3.0, facecolor="#F6F7F8",
                                     edgecolor="none", zorder=0))
    roadmap_texts = [tt for tt in c.texts if c.bbox(tt)[3] < top_panel_bottom + 1]
    for tg in (T_GATE0, T_GATE1, T_GATE2):
        c.vguide(tx(tg), y_diamond + 1.35, top_panel_bottom, avoid=roadmap_texts)

    c.check_layout("figureC1")
    return c.save(out_dir, "figureC1")


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("output")
    out.mkdir(parents=True, exist_ok=True)
    for path in (*figure1(out), *figure_c1(out)):
        print("wrote", path)
