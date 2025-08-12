# -*- coding: utf-8 -*-
# Roadmap → editable PPTX (all table columns editable, shapes independent)

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

import pandas as pd
from datetime import datetime, date
from calendar import monthrange

# -----------------------------
# 1) INPUT / OUTPUT
# -----------------------------
IN_XLSX  = r"C:\Path\to\Your\Input.xlsx"      # <- change
OUT_PPTX = r"C:\Path\to\Your\Roadmap.pptx"    # <- change

ROWS_PER_SLIDE = 30

# -----------------------------
# 2) SIZING / LAYOUT CONSTANTS
# -----------------------------
SLIDE_W, SLIDE_H = Inches(20), Inches(9)     # extra-wide slide

LEFT_PAD    = Inches(0.40)                    # left margin for whole table
RIGHT_PAD   = Inches(0.40)
TOP_ORIGIN  = Inches(1.30)                    # top start for table (space for legend)
BOTTOM_PAD  = Inches(0.50)                    # keep table within slide

SIDEBAR_W   = Inches(4.00)                    # Type+Workstream side area
TYPE_COL_W  = Inches(1.50)
WORK_COL_W  = SIDEBAR_W - TYPE_COL_W

HEADER_H    = Inches(1.00)                    # header row height
ROW_H       = Inches(0.85)                    # body row height (default)

COL_COUNT   = 12                               # months per slide
MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# guides/lines
NAVY        = RGBColor(0, 0, 128)
BLUE_HDR    = RGBColor(91, 155, 213)          # header bar
LIGHT_ROW   = RGBColor(224, 242, 255)
DARK_ROW    = RGBColor(190, 220, 240)
WHITE       = RGBColor(255, 255, 255)
BLACK       = RGBColor(0, 0, 0)
GREEN_TODAY = RGBColor(0, 176, 80)

# milestone sizes
CIRCLE_SIZE = Inches(0.30)
STAR_SIZE   = Inches(0.40)                     # you can change this anytime

status_colors = {
    "On Track":  RGBColor(0, 176, 80),
    "At Risk":   RGBColor(255, 192, 0),
    "Off Track": RGBColor(255, 0, 0),
    "Complete":  RGBColor(0, 112, 192),
    "TBC":       RGBColor(191, 191, 191),
}

SHAPE_MAP = {"Regular": MSO_SHAPE.OVAL, "Major": MSO_SHAPE.STAR_5_POINT}

# -----------------------------
# 3) DATA LOAD & ORDERING
# -----------------------------
df = pd.read_excel(IN_XLSX)

# normalize expected columns
df.rename(columns=lambda c: c.strip(), inplace=True)
assert {"Type","Workstream","Milestone Title","Milestone Date","Milestone Status","Milestone Type"} <= set(df.columns)

# parse date
df["Milestone Date"] = pd.to_datetime(df["Milestone Date"]).dt.date
df["Year"] = pd.to_datetime(df["Milestone Date"]).dt.year

# group label: "Type\n(Workstream)"
df["Group"] = df["Type"].astype(str).str.strip() + "\n(" + df["Workstream"].astype(str).str.strip() + ")"

# sort: by Type (case-insensitive), then Workstream (case-insensitive)
def norm(s): return (s or "").strip().casefold()
df["_type_key"] = df["Type"].map(norm)
df["_work_key"] = df["Workstream"].map(norm)
df = df.sort_values(by=["_type_key", "_work_key", "Milestone Date"], kind="stable").drop(columns=["_type_key","_work_key"])

# -----------------------------
# 4) UTILITIES
# -----------------------------
def add_legend(slide):
    """Single legend row across the top."""
    items = [
        ("Major Milestone", MSO_SHAPE.STAR_5_POINT, RGBColor(0,176,80)),  # green star icon, label only
        ("On Track",  MSO_SHAPE.OVAL, RGBColor(0,176,80)),
        ("At Risk",   MSO_SHAPE.OVAL, RGBColor(255,192,0)),
        ("Off Track", MSO_SHAPE.OVAL, RGBColor(255,0,0)),
        ("Complete",  MSO_SHAPE.OVAL, RGBColor(0,112,192)),
        ("TBC",       MSO_SHAPE.OVAL, RGBColor(191,191,191)),
    ]
    left = LEFT_PAD
    top  = Inches(0.50)
    slot = (SLIDE_W - LEFT_PAD - RIGHT_PAD) / len(items)

    for i, (label, shp_type, color) in enumerate(items):
        x = left + i*slot
        shp = slide.shapes.add_shape(shp_type, x, top, Inches(0.30), Inches(0.30))
        shp.fill.solid(); shp.fill.fore_color.rgb = color
        shp.line.fill.background()

        tb = slide.shapes.add_textbox(x + Inches(0.40), top, Inches(1.50), Inches(0.35))
        p = tb.text_frame.paragraphs[0]
        p.text = label
        p.font.size = Pt(14)

def build_full_editable_table(slide, year, groups, total_width, total_height):
    """
    Builds the full editable table (Type | Workstream | 12 month columns).
    Returns positioning info needed for shapes/lines.
    """
    from pptx.oxml.xmlchemy import ZeroOrOne

    row_count = len(groups)
    rows = row_count + 1
    cols = 2 + COL_COUNT

    # Month column width from remaining width after sidebar
    col_area_w = total_width - SIDEBAR_W
    col_w = col_area_w / COL_COUNT

    tbl_gf = slide.shapes.add_table(
        rows, cols,
        LEFT_PAD, TOP_ORIGIN,
        total_width, total_height
    )
    tbl = tbl_gf.table

    # column widths
    tbl.columns[0].width = TYPE_COL_W
    tbl.columns[1].width = WORK_COL_W
    for i in range(COL_COUNT):
        tbl.columns[2 + i].width = col_w

    # row heights
    tbl.rows[0].height = HEADER_H
    for r in range(1, rows):
        tbl.rows[r].height = ROW_H

    # header bar + text
    hdr = tbl.rows[0].cells
    # paint header background
    for c in range(cols):
        hdr[c].fill.solid()
        hdr[c].fill.fore_color.rgb = BLUE_HDR

    # header labels
    p = hdr[0].text_frame.paragraphs[0]; p.text = "Type";       p.font.bold=True; p.font.size=Pt(18); p.alignment=PP_ALIGN.CENTER
    p = hdr[1].text_frame.paragraphs[0]; p.text = "Workstream"; p.font.bold=True; p.font.size=Pt(18); p.alignment=PP_ALIGN.CENTER
    for i in range(COL_COUNT):
        p = hdr[2+i].text_frame.paragraphs[0]
        p.text = f"{MONTH_NAMES[i]} {year%100:02d}"
        p.font.size = Pt(14)
        p.alignment = PP_ALIGN.CENTER

    # body: Type/Workstream text + alternating fills for all columns
    for r, grp in enumerate(groups, start=1):
        type_txt, work_txt = grp.split("\n", 1)
        # left two text cells
        p = tbl.cell(r, 0).text_frame.paragraphs[0]; p.text = type_txt; p.font.size=Pt(12)
        p = tbl.cell(r, 1).text_frame.paragraphs[0]; p.text = work_txt.strip("()"); p.font.size=Pt(12)

        # alternating stripe per row for all columns
        stripe = LIGHT_ROW if (r % 2 == 1) else DARK_ROW
        for c in range(cols):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = stripe

    # geometry for month/row centers (do NOT use cell.shape)
    frame_left = tbl_gf.left
    frame_top  = tbl_gf.top

    cum_left = [frame_left]
    for c in range(cols):
        cum_left.append(cum_left[-1] + tbl.columns[c].width)

    month_cells = []
    for i in range(COL_COUNT):
        cidx = 2 + i
        left_i = cum_left[cidx]
        width_i = tbl.columns[cidx].width
        month_cells.append({"left": left_i, "width": width_i})

    first_month_left  = month_cells[0]["left"]
    last_month_right  = month_cells[-1]["left"] + month_cells[-1]["width"]

    row_cells = []
    y = frame_top + tbl.rows[0].height
    for r in range(row_count):
        h = tbl.rows[r+1].height
        row_cells.append({"top": y, "center": y + h/2, "bottom": y + h})
        y += h

    return tbl_gf, tbl, month_cells, row_cells, first_month_left, last_month_right

def draw_row_center_lines(slide, row_cells, x_left, x_right):
    """Navy thin horizontal line through center of each row (dates area)."""
    for rc in row_cells:
        y = rc["center"]
        ln = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, x_left, y, x_right, y
        )
        ln.line.fill.solid(); ln.line.fill.fore_color.rgb = NAVY
        ln.line.width = Pt(0.5)

def add_today_line_if_current_year(slide, year, month_cells, row_cells):
    """Green dotted vertical 'today' line if this slide is for the current year."""
    today = date.today()
    if today.year != year:
        return
    month_idx = today.month - 1
    days_in_m = monthrange(today.year, today.month)[1]
    day_frac  = (today.day - 1) / (days_in_m - 1)

    left = month_cells[month_idx]["left"]
    w    = month_cells[month_idx]["width"]
    xpos = left + day_frac * w

    top = row_cells[0]["top"]
    bottom = row_cells[-1]["bottom"]

    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, xpos, top, xpos, bottom)
    conn.line.fill.solid(); conn.line.fill.fore_color.rgb = GREEN_TODAY
    conn.line.width = Pt(2)
    conn.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

def plot_milestones(slide, df_page, groups, month_cells, row_cells):
    """Add star/circle shapes + labels. Shapes are independent from the table."""
    for _, row in df_page.iterrows():
        dt = row["Milestone Date"]
        year = dt.year
        m_idx = dt.month - 1
        dim = monthrange(year, dt.month)[1]
        day_frac = (dt.day - 1) / (dim - 1)

        # x within the month cell
        mleft  = month_cells[m_idx]["left"]
        mw     = month_cells[m_idx]["width"]
        x_center = mleft + day_frac * mw

        # y center of the group row
        gi = groups.index(row["Group"])
        y_center = row_cells[gi]["center"]

        # choose size and center shapes correctly by subtracting half
        is_major = str(row.get("Milestone Type","Regular")).strip().lower() == "major"
        size = STAR_SIZE if is_major else CIRCLE_SIZE
        half = size / 2.0

        shp = slide.shapes.add_shape(
            SHAPE_MAP["Major" if is_major else "Regular"],
            x_center - half, y_center - half, size, size
        )
        color = status_colors.get(row["Milestone Status"], RGBColor(128,128,128))
        shp.fill.solid(); shp.fill.fore_color.rgb = color
        shp.line.color.rgb = BLACK

        # label to the right, slightly above
        lbl = slide.shapes.add_textbox(
            x_center + half + Inches(0.10),
            y_center - half - Inches(0.05),
            Inches(2.2), Inches(0.35)
        )
        p = lbl.text_frame.paragraphs[0]
        p.text = row["Milestone Title"]
        p.font.size = Pt(12)

# -----------------------------
# 5) PRESENTATION BUILD
# -----------------------------
prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

add_title = True

for year in sorted(df["Year"].unique()):
    df_year = df[df["Year"] == year].copy()

    # ordered groups (already sorted Type → Workstream above)
    groups = df_year["Group"].drop_duplicates().tolist()

    # paginate 30 rows per slide
    pages = [groups[i:i+ROWS_PER_SLIDE] for i in range(0, len(groups), ROWS_PER_SLIDE)]
    total_pages = len(pages) if pages else 1

    for page_num, group_slice in enumerate(pages, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

        # legend (top)
        add_legend(slide)

        # figure total table area for this slide
        total_width  = SLIDE_W - LEFT_PAD - RIGHT_PAD
        total_height = SLIDE_H - TOP_ORIGIN - BOTTOM_PAD

        # build editable table and get geometry
        tbl_gf, tbl, month_cells, row_cells, x_left, x_right = build_full_editable_table(
            slide, year, group_slice, total_width, total_height
        )

        # draw navy center lines through dates area
        draw_row_center_lines(slide, row_cells, x_left, x_right)

        # data subset only for rows on this page
        sub = df_year[df_year["Group"].isin(group_slice)].copy()
        # keep rows sorted by our page's group order and by date (stable)
        order_idx = {g:i for i,g in enumerate(group_slice)}
        sub["_gidx"] = sub["Group"].map(order_idx)
        sub = sub.sort_values(by=["_gidx","Milestone Date"], kind="stable").drop(columns="_gidx")

        # plot milestones
        plot_milestones(slide, sub, group_slice, month_cells, row_cells)

        # today line for current year
        add_today_line_if_current_year(slide, year, month_cells, row_cells)

# -----------------------------
# 6) SAVE
# -----------------------------
prs.save(OUT_PPTX)
print("Saved:", OUT_PPTX)