# -*- coding: utf-8 -*-
"""
Roadmap PPT builder — full editable table + independent shapes

WHAT YOU GET PER SLIDE (per year, paginated):
- Compact legend at top
- One editable table (Type, Workstream, 12 months) that fills slide height
- Zebra body rows; header in blue; all text centered (h & v)
- Thin navy line through the middle of every body row
- (Same-year) a green dotted "today" vertical line placed by day-of-month
- Milestones (T0 -> star, T1 -> circle), placed fractionally within month by day
- Smart labels: avoid overlap; for last 4 months wrap to 3 words/line and prefer above
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

import pandas as pd
import numpy as np
import re
from datetime import datetime, date
from calendar import monthrange

# ------------------------------------------------------------
# 0) CONFIG
# ------------------------------------------------------------
IN_XLSX    = r"C:\path\to\Roadmap_Input.xlsx"          # <-- change
OUT_PPTX   = r"C:\path\to\Roadmap.pptx"                # <-- change

SLIDE_W_IN = 20.0
SLIDE_H_IN = 9.0

LEFT_PAD_IN   = 0.09
RIGHT_PAD_IN  = 0.09
TOP_PAD_IN    = 0.25          # small gap above legend
LEGEND_H_IN   = 0.45          # more compact legend
HEADER_H_IN   = 1.0           # months header height
BOTTOM_PAD_IN = 0.09

TYPE_COL_W_IN = 1.6
WORK_COL_W_IN = 2.8
MAX_ROWS_PER_SLIDE = 20

# shape sizes
CIRCLE_D_IN = 0.30
STAR_D_IN   = 0.40
LABEL_W_IN  = 2.4         # default label width
LABEL_H_IN  = 0.45

# colors
BLUE_HDR   = RGBColor(91,155,213)
WHITE      = RGBColor(255,255,255)
TEXT_DARK  = RGBColor(0,0,0)
MONTH_ODD  = RGBColor(224,242,255)
MONTH_EVEN = RGBColor(190,220,240)
NAVY       = RGBColor(0,0,128)
GREEN_TOD  = RGBColor(0,176,80)

STATUS_COLORS = {
    "on track":   RGBColor(0,176,80),
    "at risk":    RGBColor(255,192,0),
    "off track":  RGBColor(255,0,0),
    "complete":   RGBColor(0,112,192),
    "tbc":        RGBColor(191,191,191),
}

# ------------------------------------------------------------
# 1) HELPERS: cleaning, sorting, grouping
# ------------------------------------------------------------
def clean(s):
    if pd.isna(s): return ""
    return str(s).strip().replace("\n"," ").casefold()

def type_bucket(s):
    s = clean(s)
    m = re.match(r"([a-z0-9]+)", s)
    return m.group(1) if m else s

def build_groups_for_year(df_year_sorted):
    """
    Return an ordered list of unique "Type\nWorkstream" pairs following
    the stable sorted order of df_year_sorted.
    """
    # make display label exactly "Type\nWorkstream" (original case)
    grp_series = df_year_sorted.apply(lambda r: f"{r['Type']}\n{r['Workstream']}", axis=1)
    groups = pd.Series(grp_series).drop_duplicates().tolist()
    return groups

def slice_pages(groups, max_rows=MAX_ROWS_PER_SLIDE):
    for i in range(0, len(groups), max_rows):
        yield i//max_rows + 1, groups[i:i+max_rows]

# ------------------------------------------------------------
# 2) TABLE BUILDER (full editable grid)
# ------------------------------------------------------------
def build_full_table(slide, groups, year):
    """
    Creates one editable table (2 + 12 columns) that fills the available
    vertical space down to BOTTOM_PAD_IN and returns geometry for drawing.
    """
    rows = int(len(groups) + 1)  # header + body
    cols = int(14)               # Type, Workstream, 12 months

    left_in  = LEFT_PAD_IN
    right_in = SLIDE_W_IN - RIGHT_PAD_IN
    total_w  = right_in - left_in

    # columns
    type_w = TYPE_COL_W_IN
    work_w = WORK_COL_W_IN
    months_w = max(0.01, total_w - (type_w + work_w))
    month_w  = months_w / 12.0

    # heights – fill all the way to bottom pad
    top_in = TOP_PAD_IN + LEGEND_H_IN + 0.10  # little gap below legend
    avail_h = SLIDE_H_IN - top_in - BOTTOM_PAD_IN
    header_h = HEADER_H_IN
    body_h   = max(0.01, avail_h - header_h)
    body_rows = max(1, rows-1)
    row_h = body_h / body_rows

    # create
    tbl_shape = slide.shapes.add_table(
        rows, cols,
        Inches(left_in),
        Inches(top_in),
        Inches(total_w),
        Inches(header_h + row_h*(rows-1))
    )
    tbl = tbl_shape.table
    # style supplies grid borders (avoids 'Cell has no border_left' error)
    tbl.style = 'Table Grid'

    # column widths (EMU via Inches ensures integers)
    tbl.columns[0].width = Inches(type_w)
    tbl.columns[1].width = Inches(work_w)
    for c in range(12):
        tbl.columns[2+c].width = Inches(month_w)

    # row heights
    tbl.rows[0].height = Inches(header_h)
    for r in range(1, rows):
        tbl.rows[r].height = Inches(row_h)

    # header fill + text
    hdr_cells = tbl.rows[0].cells
    hdr_cells[0].text = "Type"
    hdr_cells[1].text = "Workstream"
    # months
    for m_idx in range(12):
        dt = date(year, m_idx+1, 1)
        hdr_cells[2+m_idx].text = dt.strftime("%b %y")

    # center header text (white) and blue fill
    for c in range(cols):
        p = hdr_cells[c].text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        hdr_cells[c].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        for run in p.runs: run.font.size = Pt(18)
        hdr_cells[c].fill.solid()
        hdr_cells[c].fill.fore_color.rgb = BLUE_HDR
        # make header text white
        for run in p.runs: run.font.color.rgb = WHITE

    # body zebra + center alignment + put Type/Workstream values
    for r, grp in enumerate(groups, start=1):
        fill_rgb = MONTH_ODD if (r % 2 == 1) else MONTH_EVEN
        for c in range(cols):
            cell = tbl.rows[r].cells[c]
            cell.fill.solid()
            cell.fill.fore_color.rgb = fill_rgb
            cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(15)
            p.font.color.rgb = TEXT_DARK

        # split "Type\nWorkstream"
        t, w = grp.split("\n", 1)
        tbl.rows[r].cells[0].text = t
        tbl.rows[r].cells[1].text = w

    # ----- geometry we’ll need for drawing shapes -----
    left_months  = left_in + type_w + work_w
    right_months = left_months + months_w
    y_tops   = [ (top_in + header_h + row_h*(r-1)) for r in range(1, rows) ]
    y_centers = [ y + row_h/2.0 for y in y_tops ]  # one per body row

    geom = {
        "left_in": left_in,
        "right_in": right_in,
        "top_in": top_in,
        "header_h": header_h,
        "row_h": row_h,
        "left_months_in": left_months,
        "right_months_in": right_months,
        "month_w_in": month_w,
        "y_centers_in": y_centers,  # aligned with groups order
    }
    return tbl_shape, tbl, geom

# ------------------------------------------------------------
# 3) NAVY center lines (thin), today line (green dotted)
# ------------------------------------------------------------
def draw_row_center_lines(slide, geom, row_count):
    x1 = Inches(geom["left_months_in"])
    x2 = Inches(geom["right_months_in"])
    for y_in in geom["y_centers_in"][:row_count]:
        y = Inches(y_in)
        ln = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, x1, y, x2, y
        )
        ln.line.fill.solid()
        ln.line.fill.fore_color.rgb = NAVY
        ln.line.width = Pt(0.5)

def add_today_line_if_same_year(slide, year, geom):
    today = date.today()
    if today.year != year:
        return
    days_in = monthrange(today.year, today.month)[1]
    month_idx = today.month - 1
    # spread by real day (no mid-month centering)
    day_frac = (today.day - 1) / days_in
    xpos_in = geom["left_months_in"] + (month_idx + day_frac) * geom["month_w_in"]

    ln = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(xpos_in), Inches(geom["top_in"]),
        Inches(xpos_in), Inches(geom["top_in"] + geom["row_h"]*len(geom["y_centers_in"]))
    )
    ln.line.fill.solid()
    ln.line.fill.fore_color.rgb = GREEN_TOD
    ln.line.width = Pt(2)
    ln.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

# ------------------------------------------------------------
# 4) Milestones & labels
# ------------------------------------------------------------
def three_word_wrap(text):
    words = text.split()
    out, line = [], []
    for w in words:
        line.append(w)
        if len(line) == 3:
            out.append(" ".join(line))
            line = []
    if line:
        out.append(" ".join(line))
    return "\n".join(out)

def place_labels_nonoverlap(slide, base_x_in, base_y_in, text, prefer_above,
                            months_right_in, existing_rects, max_try=8):
    """
    Try to place a small textbox near (base_x_in, base_y_in) without
    overlapping previously placed labels in the same row.
    """
    width_in = LABEL_W_IN
    height_in = LABEL_H_IN
    # if too near right edge, prefer above
    if base_x_in + width_in > months_right_in:
        prefer_above = True

    for i in range(max_try):
        dy_in = (0.18 + 0.18*i) * ( -1 if prefer_above or (i%2==0) else 1 )
        top_in = base_y_in + dy_in - height_in/2.0
        left_in = min(base_x_in + 0.1, months_right_in - width_in)  # keep within months area

        rect = (left_in, top_in, left_in+width_in, top_in+height_in)
        # overlap check
        overlaps = False
        for (x1,y1,x2,y2) in existing_rects:
            if not (rect[2] <= x1 or rect[0] >= x2 or rect[3] <= y1 or rect[1] >= y2):
                overlaps = True
                break
        if not overlaps:
            tb = slide.shapes.add_textbox(Inches(left_in), Inches(top_in),
                                          Inches(width_in), Inches(height_in))
            tf = tb.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.text = text
            p.alignment = PP_ALIGN.LEFT
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_DARK
            existing_rects.append(rect)
            return

    # fallback: just drop a minimal offset above
    tb = slide.shapes.add_textbox(Inches(base_x_in+0.1), Inches(base_y_in-0.25),
                                  Inches(width_in), Inches(height_in))
    tf = tb.text_frame
    tf.paragraphs[0].text = text
    tf.paragraphs[0].font.size = Pt(12)
    tf.paragraphs[0].font.color.rgb = TEXT_DARK
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

def plot_milestones(slide, df_page, groups, geom, year):
    left_m = geom["left_months_in"]
    month_w = geom["month_w_in"]
    months_right = geom["right_months_in"]
    y_centers = geom["y_centers_in"]

    # per-row placed label rects
    row_label_rects = {i: [] for i in range(len(groups))}

    for _, row in df_page.iterrows():
        # row index on this page
        grp = f"{row['Type']}\n{row['Workstream']}"
        try:
            ri = groups.index(grp)
        except ValueError:
            continue

        dt: pd.Timestamp = row["Milestone Date"]
        m_idx = dt.month - 1
        days_in = monthrange(dt.year, dt.month)[1]
        frac = (dt.day - 1) / days_in

        x_in = left_m + (m_idx + frac) * month_w
        y_in = y_centers[ri]

        # choose shape by type (T0 star, T1 circle), case-insensitive
        mt = clean(row.get("Milestone Type", ""))
        shp_size = STAR_D_IN if mt in ("t0","major") else CIRCLE_D_IN
        shp_kind = MSO_SHAPE.STAR_5_POINT if mt in ("t0","major") else MSO_SHAPE.OVAL

        shp = slide.shapes.add_shape(shp_kind,
                                     Inches(x_in - shp_size/2.0),
                                     Inches(y_in - shp_size/2.0),
                                     Inches(shp_size), Inches(shp_size))
        shp.fill.solid()
        status = clean(row.get("Milestone Status",""))
        shp.fill.fore_color.rgb = STATUS_COLORS.get(status, RGBColor(0,176,80))
        shp.line.color.rgb = RGBColor(0,0,0)

        # label text rules
        title = str(row.get("Milestone Title","")).strip()
        prefer_above = (m_idx >= 8)  # last 4 months bias above
        if m_idx >= 8:   # also wrap to 3 words per line in last 4 months
            title = three_word_wrap(title)

        place_labels_nonoverlap(
            slide,
            base_x_in = x_in + shp_size/2.0,
            base_y_in = y_in,
            text = title,
            prefer_above = prefer_above,
            months_right_in = months_right,
            existing_rects = row_label_rects[ri]
        )

# ------------------------------------------------------------
# 5) Legend (compact)
# ------------------------------------------------------------
def add_legend(slide, left_in=LEFT_PAD_IN, top_in=TOP_PAD_IN, height_in=LEGEND_H_IN):
    items = [
        ("T0", MSO_SHAPE.STAR_5_POINT, STATUS_COLORS["on track"]),
        ("T1", MSO_SHAPE.OVAL, STATUS_COLORS["on track"]),
        ("On Track", MSO_SHAPE.OVAL, STATUS_COLORS["on track"]),
        ("At Risk",  MSO_SHAPE.OVAL, STATUS_COLORS["at risk"]),
        ("Off Track",MSO_SHAPE.OVAL, STATUS_COLORS["off track"]),
        ("Complete", MSO_SHAPE.OVAL, STATUS_COLORS["complete"]),
        ("TBC",      MSO_SHAPE.OVAL, STATUS_COLORS["tbc"]),
    ]
    slot_w = (SLIDE_W_IN - LEFT_PAD_IN - RIGHT_PAD_IN) / len(items)
    cy = top_in + height_in/2.0
    for i,(lbl, shp_kind, col) in enumerate(items):
        cx = left_in + (i+0.5)*slot_w
        # shape
        s = slide.shapes.add_shape(
            shp_kind,
            Inches(cx - 0.12), Inches(cy - 0.12),
            Inches(0.24), Inches(0.24)
        )
        s.fill.solid(); s.fill.fore_color.rgb = col
        # label
        tb = slide.shapes.add_textbox(Inches(cx - 0.12 + 0.3), Inches(cy - 0.14),
                                      Inches(1.2), Inches(0.28))
        tf = tb.text_frame; tf.clear()
        p = tf.paragraphs[0]; p.text = lbl; p.font.size = Pt(12)

# ------------------------------------------------------------
# 6) Slide builder (per year per page)
# ------------------------------------------------------------
def build_slide(prs, df_page, year, groups, page_no, total_pages):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    add_legend(slide)  # compact legend

    # full editable table + geometry
    tbl_shape, tbl, geom = build_full_table(slide, groups, year)

    # navy center lines
    draw_row_center_lines(slide, geom, len(groups))

    # "today" only if same year
    add_today_line_if_same_year(slide, year, geom)

    # plot milestones for this page only
    plot_milestones(slide, df_page, groups, geom, year)

# ------------------------------------------------------------
# 7) MAIN
# ------------------------------------------------------------
def main():
    df = pd.read_excel(IN_XLSX)

    # required columns (case-sensitive names expected in your sheet)
    # Type, Workstream, Milestone Title, Milestone Date, Milestone Type (T0/T1), Milestone Status
    df["Milestone Date"] = pd.to_datetime(df["Milestone Date"])
    df["year"] = df["Milestone Date"].dt.year

    # sorting helpers
    df["Type_key"] = df["Type"].map(clean)
    df["Type_bucket"] = df["Type"].map(type_bucket)
    df["Work_key"] = df["Workstream"].map(clean)

    # global stable sort: bucket -> type -> work -> date
    df_sorted = df.sort_values(
        by=["Type_bucket", "Type_key", "Work_key", "Milestone Date"],
        kind="stable"
    )

    prs = Presentation()
    prs.slide_width  = Inches(SLIDE_W_IN)
    prs.slide_height = Inches(SLIDE_H_IN)

    years = sorted(df_sorted["year"].unique().tolist())
    for year in years:
        df_year = df_sorted[df_sorted["year"] == year].copy()

        # ordered groups for this year, following sorted order
        groups = build_groups_for_year(df_year)

        # paginate by groups (20 per slide)
        pages = list(slice_pages(groups, MAX_ROWS_PER_SLIDE))
        total_pages = len(pages)

        for page_no, grp_slice in pages:
            # rows for this page only, keep order
            df_page = df_year[df_year.apply(
                lambda r: f"{r['Type']}\n{r['Workstream']}" in set(grp_slice), axis=1
            )].copy()

            # build the slide
            build_slide(prs, df_page, year, grp_slice, page_no, total_pages)

    prs.save(OUT_PPTX)
    print("Saved:", OUT_PPTX)

if __name__ == "__main__":
    main()