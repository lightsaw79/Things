# -*- coding: utf-8 -*-
# Roadmap -> PPTX (editable table + independent milestones)
#
# Option A: true borders on every table cell via OXML (no AttributeError).
#
# --- Requirements you asked for ---
# - Wide slide, 12 month columns + 2 left columns (Type, Workstream)
# - Editable table (all columns resizable); thin white border on each cell
# - Alternating row fill across the whole width
# - Navy center line per row across the dates area
# - Accurate milestone x-position (by day-of-month), y centered to row
# - Shapes independent of table resizing (they won't move if you drag columns)
# - Full-width legend above the table
# - Green dotted 'today' line on the current year only
# - Year-based slides; paginate 30 rows per slide
#
# ---------------------------------------------------------------------

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

import pandas as pd
import re
from datetime import datetime, date
from calendar import monthrange

# --------------------------
# CONFIG / TUNING CONSTANTS
# --------------------------
IN_X = Inches  # shorthand

SLIDE_W_IN = 20
SLIDE_H_IN = 9

TOP_PAD_IN     = 0.40   # gap above legend
LEGEND_H_IN    = 0.60
LEGEND_TEXT_PT = 15
GAP_AFTER_LEGEND_IN = 0.20

LEFT_PAD_IN    = 0.40   # outside left margin
RIGHT_PAD_IN   = 0.40   # outside right margin
BOTTOM_PAD_IN  = 0.50   # outside bottom margin

TYPE_COL_W_IN  = 1.50
WORK_COL_W_IN  = 2.50
HEADER_H_IN    = 1.00
ROW_H_IN       = 0.95

COLS_TOTAL = 14  # Type, Workstream, Jan..Dec (12 months)

# Colors
BLUE_HDR   = RGBColor( 91,155,213)
MONTH_ODD  = RGBColor(224,242,255)
MONTH_EVEN = RGBColor(190,220,240)
WHITE      = RGBColor(255,255,255)
NAVY       = RGBColor(  0,  0,128)
TODAY_GRN  = RGBColor(  0,176, 80)

# Milestone sizes
CIRCLE_SIZE_IN = 0.30
STAR_SIZE_IN   = 0.40  # Major milestone

# Legend items (full-width)
LEGEND = [
    ("Major Milestone", MSO_SHAPE.STAR_5_POINT, RGBColor(  0,176, 80)),
    ("On Track",        MSO_SHAPE.OVAL,        RGBColor(  0,176, 80)),
    ("At Risk",         MSO_SHAPE.OVAL,        RGBColor(255,192,  0)),
    ("Off Track",       MSO_SHAPE.OVAL,        RGBColor(255,  0,  0)),
    ("Complete",        MSO_SHAPE.OVAL,        RGBColor(  0,112,192)),
    ("TBC",             MSO_SHAPE.OVAL,        RGBColor(191,191,191)),
]

STATUS_COLORS = {
    "On Track":  RGBColor(  0,176, 80),
    "At Risk":   RGBColor(255,192,  0),
    "Off Track": RGBColor(255,  0,  0),
    "Complete":  RGBColor(  0,112,192),
    "TBC":       RGBColor(191,191,191),
}

# -------------
# UTIL HELPERS
# -------------
def set_cell_border(cell, color_hex="FFFFFF", width="6350"):
    """
    Apply solid borders to a python-pptx table cell using OXML.
    color_hex: 'RRGGBB' (no '#')
    width:     an EMU-ish width value string; 6350 ≈ thin, 12700 thicker
    """
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()

    for edge in ("L", "R", "T", "B"):
        ln = OxmlElement(f'a:ln{edge}')
        ln.set('w', width)

        solid = OxmlElement('a:solidFill')
        rgb   = OxmlElement('a:srgbClr')
        rgb.set('val', color_hex)
        solid.append(rgb)
        ln.append(solid)

        dash = OxmlElement('a:prstDash')
        dash.set('val', 'solid')
        ln.append(dash)

        tcPr.append(ln)

def clean_text(s: str) -> str:
    return str(s).strip().replace("\n", " ").casefold()

def first_token_or_all(s: str) -> str:
    s = clean_text(s)
    m = re.match(r"([a-z0-9]+)", re.sub(r"[^a-z0-9]", " ", s))
    return m.group(1) if m else s

def month_name(m_idx: int, year: int) -> str:
    dt = date(year, m_idx+1, 1)
    return dt.strftime("%b %y")

# ------------------------------------------------------
# GEOMETRY: build the full editable table and return
#           useful positions to place lines and shapes
# ------------------------------------------------------
def build_full_editable_table(slide, year, groups, total_w_in, total_h_in):
    """
    Returns:
      tbl           : the pptx table object
      month_left_in : left x (inches) of the Jan column (inside table)
      month_w_in    : width (inches) of one month column
      first_row_top_in : top y (inches) of the first body row
    """

    left_origin_in = LEFT_PAD_IN
    top_origin_in  = TOP_PAD_IN + LEGEND_H_IN + GAP_AFTER_LEGEND_IN

    # column widths
    dates_width_in = total_w_in - TYPE_COL_W_IN - WORK_COL_W_IN
    month_w_in     = dates_width_in / 12.0

    row_count = len(groups)

    # create table: header row + body rows
    rows = int(row_count + 1)
    cols = int(14)

    tbl_shape = slide.shapes.add_table(
        rows, cols,
        IN_X(left_origin_in), IN_X(top_origin_in),
        IN_X(total_w_in), IN_X(total_h_in)
    )
    tbl = tbl_shape.table

    # set column widths
    tbl.columns[0].width = IN_X(TYPE_COL_W_IN)
    tbl.columns[1].width = IN_X(WORK_COL_W_IN)
    for c in range(12):
        tbl.columns[2+c].width = IN_X(month_w_in)

    # header height + body heights
    tbl.rows[0].height = IN_X(HEADER_H_IN)
    for r in range(1, rows):
        tbl.rows[r].height = IN_X(ROW_H_IN)

    # header cells
    hdr_type = tbl.cell(0,0)
    hdr_work = tbl.cell(0,1)
    for cell, title in ((hdr_type,"Type"), (hdr_work,"Workstream")):
        cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_HDR
        p = cell.text_frame.paragraphs[0]
        p.text = title
        p.font.bold = True
        p.font.size = Pt(18)
        p.font.color.rgb = WHITE
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p.alignment = PP_ALIGN.CENTER
        set_cell_border(cell, "FFFFFF", "6350")

    # month headers
    for m in range(12):
        hc = tbl.cell(0, 2+m)
        hc.fill.solid(); hc.fill.fore_color.rgb = BLUE_HDR
        p = hc.text_frame.paragraphs[0]
        p.text = month_name(m, year)
        p.font.bold = True
        p.font.size = Pt(18)
        p.font.color.rgb = WHITE
        hc.vertical_anchor = MSO_ANCHOR.MIDDLE
        p.alignment = PP_ALIGN.CENTER
        set_cell_border(hc, "FFFFFF", "6350")

    # body rows (left labels + alternating month fills + borders)
    for r, grp in enumerate(groups, start=1):
        type_txt, work_txt = grp.split("\n", 1)

        # Type cell
        c0 = tbl.cell(r, 0)
        c0.fill.solid(); c0.fill.fore_color.rgb = MONTH_ODD if (r % 2 == 1) else MONTH_EVEN
        p0 = c0.text_frame.paragraphs[0]
        p0.text = type_txt
        p0.font.size = Pt(15)
        p0.font.color.rgb = RGBColor(0,0,0)
        c0.vertical_anchor = MSO_ANCHOR.MIDDLE
        p0.alignment = PP_ALIGN.CENTER
        set_cell_border(c0, "FFFFFF", "6350")

        # Workstream cell
        c1 = tbl.cell(r, 1)
        c1.fill.solid(); c1.fill.fore_color.rgb = MONTH_ODD if (r % 2 == 1) else MONTH_EVEN
        p1 = c1.text_frame.paragraphs[0]
        p1.text = work_txt
        p1.font.size = Pt(15)
        p1.font.color.rgb = RGBColor(0,0,0)
        c1.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1.alignment = PP_ALIGN.CENTER
        set_cell_border(c1, "FFFFFF", "6350")

        # Month cells
        for c in range(12):
            cc = tbl.cell(r, 2+c)
            cc.fill.solid()
            cc.fill.fore_color.rgb = MONTH_ODD if (r % 2 == 1) else MONTH_EVEN
            set_cell_border(cc, "FFFFFF", "6350")

    # geometry for later overlays
    first_row_top_in = top_origin_in + HEADER_H_IN
    month_left_in    = left_origin_in + TYPE_COL_W_IN + WORK_COL_W_IN

    return tbl, month_left_in, month_w_in, first_row_top_in

# -------------------------
# LEGEND (full-width)
# -------------------------
def add_full_width_legend(slide, left_in, top_in, width_in, height_in):
    slot_w = width_in / len(LEGEND)
    y      = top_in + height_in/2.0

    for i, (label, shp_type, color) in enumerate(LEGEND):
        cx = left_in + i*slot_w + slot_w*0.5
        s  = slide.shapes.add_shape(shp_type,
                                    IN_X(cx - 0.15), IN_X(y - 0.15),
                                    IN_X(0.30), IN_X(0.30))
        s.fill.solid(); s.fill.fore_color.rgb = color
        s.line.fill.background()

        tb = slide.shapes.add_textbox(IN_X(cx + 0.05), IN_X(y - 0.18), IN_X(slot_w*0.7), IN_X(0.36))
        p  = tb.text_frame.paragraphs[0]
        p.text = label
        p.font.size = Pt(LEGEND_TEXT_PT)
        p.alignment = PP_ALIGN.LEFT

# -----------------------------------
# Navy row-center lines in dates area
# -----------------------------------
def draw_row_center_lines(slide, row_count, first_row_top_in, month_left_in, month_w_in, table_width_in):
    dates_left  = month_left_in
    dates_right = month_left_in + 12*month_w_in
    for r in range(row_count):
        y = first_row_top_in + r*ROW_H_IN + ROW_H_IN/2.0
        ln = slide.shapes.add_shape(
            MSO_CONNECTOR.STRAIGHT,
            IN_X(dates_left), IN_X(y),
            IN_X(dates_right - dates_left), IN_X(0)
        )
        ln.line.fill.solid()
        ln.line.fill.fore_color.rgb = NAVY
        ln.line.width = Pt(0.5)

# -----------------------------------
# Milestones (independent shapes)
# -----------------------------------
def plot_milestones(slide, df_page, groups, year, month_left_in, month_w_in, first_row_top_in):
    for _, row in df_page.iterrows():
        dt: date = row["Milestone Date"].date() if isinstance(row["Milestone Date"], pd.Timestamp) else row["Milestone Date"]
        if dt.year != year:  # safety
            continue

        month_idx   = dt.month - 1
        days_in_mon = monthrange(dt.year, dt.month)[1]
        day_frac    = (dt.day - 1) / (days_in_mon - 1 if days_in_mon > 1 else 1)

        # Row index
        grp_label = f"{row['Type']}\n{row['Workstream']}"
        try:
            r = groups.index(grp_label)
        except ValueError:
            continue

        # Shape and size
        major = str(row.get("Milestone Type","")).strip().lower() == "major"
        size  = STAR_SIZE_IN if major else CIRCLE_SIZE_IN

        # Position (centered)
        x_center = month_left_in + month_idx*month_w_in + day_frac*month_w_in
        y_center = first_row_top_in + r*ROW_H_IN + ROW_H_IN/2.0

        shp = slide.shapes.add_shape(
            MSO_SHAPE.STAR_5_POINT if major else MSO_SHAPE.OVAL,
            IN_X(x_center - size/2.0), IN_X(y_center - size/2.0),
            IN_X(size), IN_X(size)
        )
        shp.fill.solid()
        shp.fill.fore_color.rgb = STATUS_COLORS.get(str(row.get("Milestone Status","")).strip(), RGBColor(128,128,128))
        shp.line.color.rgb = RGBColor(0,0,0)

        # Label (right & a little above)
        lbl = slide.shapes.add_textbox(
            IN_X(x_center + size*0.55), IN_X(y_center - size*0.70),
            IN_X(2.5), IN_X(0.35)
        )
        p = lbl.text_frame.paragraphs[0]
        p.text = str(row.get("Milestone Title","")).strip()
        p.font.size = Pt(12)
        p.alignment = PP_ALIGN.LEFT

# -----------------------------------
# Today's dotted vertical line
# -----------------------------------
def draw_today_line(slide, year, month_left_in, month_w_in, first_row_top_in, row_count):
    today = date.today()
    if today.year != year:
        return
    days_in = monthrange(today.year, today.month)[1]
    month_idx = today.month - 1
    day_frac  = (today.day - 1) / (days_in - 1 if days_in > 1 else 1)
    x = month_left_in + (month_idx + day_frac) * month_w_in

    top = TOP_PAD_IN + LEGEND_H_IN + GAP_AFTER_LEGEND_IN + HEADER_H_IN
    height = row_count * ROW_H_IN
    conn = slide.shapes.add_shape(
        MSO_CONNECTOR.STRAIGHT,
        IN_X(x), IN_X(top),
        IN_X(0), IN_X(height)
    )
    conn.line.fill.solid()
    conn.line.fill.fore_color.rgb = TODAY_GRN
    conn.line.width = Pt(2)
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    conn.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

# -----------------------------------
# Build one slide from a page (<=30)
# -----------------------------------
def build_slide(prs, df_page, year, page_no, total_pages):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Title (optional – tweak as you like)
    title = slide.shapes.add_textbox(IN_X(LEFT_PAD_IN), IN_X(0.1), IN_X(10), IN_X(0.5))
    tp = title.text_frame.paragraphs[0]
    tp.text = f"TM‑US Roadmap — {year}" + (f" (Page {page_no}/{total_pages})" if total_pages > 1 else "")
    tp.font.size = Pt(22); tp.font.bold = True

    # Legend
    content_left_in  = LEFT_PAD_IN
    content_width_in = SLIDE_W_IN - LEFT_PAD_IN - RIGHT_PAD_IN
    add_full_width_legend(slide, content_left_in, TOP_PAD_IN, content_width_in, LEGEND_H_IN)

    # Groups (y‑axis labels)
    groups = (df_page[["Type","Workstream"]]
              .apply(lambda s: f"{s['Type']}\n{s['Workstream']}", axis=1)
              .tolist())

    # Table geometry
    total_height_in = SLIDE_H_IN - (TOP_PAD_IN + LEGEND_H_IN + GAP_AFTER_LEGEND_IN) - BOTTOM_PAD_IN
    tbl, month_left_in, month_w_in, first_row_top_in = build_full_editable_table(
        slide, year, groups, content_width_in, total_height_in
    )

    # Navy row center lines
    draw_row_center_lines(slide, len(groups), first_row_top_in, month_left_in, month_w_in, content_width_in)

    # Milestones
    plot_milestones(slide, df_page, groups, year, month_left_in, month_w_in, first_row_top_in)

    # Today line
    draw_today_line(slide, year, month_left_in, month_w_in, first_row_top_in, len(groups))

# -----------------------------------
# MAIN
# -----------------------------------
def main(input_xlsx, out_pptx):
    df = pd.read_excel(input_xlsx)

    # Ensure datetime
    df["Milestone Date"] = pd.to_datetime(df["Milestone Date"])
    df["year"] = df["Milestone Date"].dt.year

    # Sorting helpers (bucket + cleaned keys)
    df["Type_key"] = df["Type"].map(clean_text)
    df["Type_bucket"] = df["Type"].map(first_token_or_all)
    df["Work_key"] = df["Workstream"].map(clean_text)

    # Global sort so pages come out consistent
    df = df.sort_values(by=["Type_bucket", "Type_key", "Work_key", "Milestone Date"], kind="stable")

    prs = Presentation()
    prs.slide_width  = IN_X(SLIDE_W_IN)
    prs.slide_height = IN_X(SLIDE_H_IN)

    for year in sorted(df["year"].dropna().unique()):
        df_year = df[df["year"] == year].copy()

        # Page order for groups (unique, in sorted order)
        groups_order = (df_year[["Type_key","Type_bucket","Work_key","Type","Workstream"]]
                        .drop_duplicates()
                        .sort_values(by=["Type_bucket","Type_key","Work_key"], kind="stable"))

        # Build the label column we’ll use for grouping/page slicing
        df_year["Group"] = df_year.apply(lambda s: f"{s['Type']}\n{s['Workstream']}", axis=1)

        # Pagination: 30 rows per slide
        groups = groups_order.apply(lambda s: f"{s['Type']}\n{s['Workstream']}", axis=1).tolist()
        total_rows = len(groups)
        pages = [groups[i:i+30] for i in range(0, total_rows, 30)]
        total_pages = len(pages) if pages else 1

        if not pages:
            continue

        for page_no, group_slice in enumerate(pages, start=1):
            # Take only rows whose Group in this slice, and sort by that slice order then by date
            order_index = {g:i for i,g in enumerate(group_slice)}
            sub = df_year[df_year["Group"].isin(group_slice)].copy()
            sub["_gidx"] = sub["Group"].map(order_index)
            sub = sub.sort_values(by=["_gidx","Milestone Date"], kind="stable").drop(columns="_gidx")

            build_slide(prs, sub, year, page_no, total_pages)

    prs.save(out_pptx)
    print("Saved:", out_pptx)


# ----------------------
# Run it (edit paths)
# ----------------------
if __name__ == "__main__":
    INPUT_XLSX = r"C:\Users\YOU\Desktop\Roadmap_Input_Sheet.xlsx"  # <-- change me
    OUTPUT_PPTX = r"C:\Users\YOU\Desktop\Roadmap.pptx"             # <-- change me
    main(INPUT_XLSX, OUTPUT_PPTX)