# -*- coding: utf-8 -*-
"""
Roadmap -> Editable PPT (python-pptx)
- Left sidebar is a real table (resizable columns)
- Date grid is shapes (rectangles) with alternating shading + borders
- Milestones placed by exact day-of-month
- Slides are partitioned by Year; each Year paginates 30 rows/slide
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

import pandas as pd
from datetime import date as date_cls, datetime
from calendar import monthrange

# -------------------------------------------------------------------
# 1) CONFIG
# -------------------------------------------------------------------

INPUT_XLSX = r"C:\Users\you\Desktop\Roadmap_Input.xlsx"   # <-- change
OUT_PPTX   = r"C:\Users\you\Desktop\Roadmap.pptx"         # <-- change

ROWS_PER_SLIDE = 30

# Colors
BLUE_HDR   = RGBColor( 91,155,213)  # month header & top bar
COL_LIGHT  = RGBColor(224,242,255)  # alternating grid light
COL_DARK   = RGBColor(190,220,240)  # alternating grid dark
WHITE      = RGBColor(255,255,255)
BLACK      = RGBColor(  0,  0,  0)
NAVY       = RGBColor(  0,  0,128)
GREEN_DOT  = RGBColor(  0,176, 80)

# Milestone status color map
STATUS_COLORS = {
    "On Track" : RGBColor(  0,176, 80),  # green
    "At Risk"  : RGBColor(255,192,  0),  # yellow
    "Off Track": RGBColor(255,  0,  0),  # red
    "Complete" : RGBColor(  0,112,192),  # blue
    "TBC"      : RGBColor(191,191,191),  # gray
}

# Milestone sizes
CIRCLE_SIZE = Inches(0.30)  # regular
STAR_SIZE   = Inches(0.40)  # major

# Layout (all values in EMUs; Inches() returns EMU int)
SLIDE_W   = Inches(20)        # extra wide
SLIDE_H   = Inches(9)

LEFT_PAD  = Inches(0.25)
RIGHT_PAD = Inches(0.25)

SIDEBAR_W   = Inches(4.0)     # left table area
TYPE_COL_W  = Inches(1.5)
WORK_COL_W  = SIDEBAR_W - TYPE_COL_W

HEADER_H     = Inches(1.0)    # month header row height
LEGEND_H     = Inches(0.60)
LEGEND_TOP   = Inches(0.20)   # top padding before legend
TOP_MARGIN   = LEGEND_TOP + LEGEND_H + Inches(0.20)
BOTTOM_MARGIN= Inches(0.50)

# -------------------------------------------------------------------
# 2) LOAD & PREP DATA
# -------------------------------------------------------------------

df = pd.read_excel(INPUT_XLSX)
df["Milestone Date"] = pd.to_datetime(df["Milestone Date"])
df["Year"] = df["Milestone Date"].dt.year
df["Group"] = df["Type"].astype(str) + "\n(" + df["Workstream"].astype(str) + ")"

# Sort groups: by Type bucket (case-insensitive cleaned) then Workstream alpha
def clean_key(s: str) -> str:
    return (str(s).strip().casefold()
            .replace("_"," ").replace("-"," ")
            .replace("  "," "))

df["Type_key"] = df["Type"].map(clean_key)
df["Work_key"] = df["Workstream"].map(clean_key)

df = df.sort_values(by=["Type_key", "Work_key", "Milestone Date"], kind="stable")

# -------------------------------------------------------------------
# 3) HELPER: build one slide (subset_df) with up to N rows
# -------------------------------------------------------------------

def build_slide(prs: Presentation, subset_df: pd.DataFrame, year: int, page_num: int, total_pages: int):
    """Render one slide for 'year' with the rows in subset_df (<= ROWS_PER_SLIDE)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # set slide size on presentation if not set
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    # compute chart area (in EMUs ints)
    chart_w = SLIDE_W - SIDEBAR_W - RIGHT_PAD
    chart_h = SLIDE_H - TOP_MARGIN - BOTTOM_MARGIN

    groups = (subset_df["Type"] + "\n(" + subset_df["Workstream"] + ")").tolist()
    row_count = len(groups)
    col_count = 12

    col_w = chart_w // col_count
    row_h = chart_h // max(1, row_count)

    left_origin = LEFT_PAD         # left of sidebar
    top_origin  = TOP_MARGIN

    # ----- 3.1 Legend spanning entire top -----
    legend_items = [
        ("Major Milestone", MSO_SHAPE.STAR_5_POINT, RGBColor(  0,176, 80)),  # just an icon sample color
        ("On Track",        MSO_SHAPE.OVAL,         STATUS_COLORS["On Track"]),
        ("At Risk",         MSO_SHAPE.OVAL,         STATUS_COLORS["At Risk"]),
        ("Off Track",       MSO_SHAPE.OVAL,         STATUS_COLORS["Off Track"]),
        ("Complete",        MSO_SHAPE.OVAL,         STATUS_COLORS["Complete"]),
        ("TBC",             MSO_SHAPE.OVAL,         STATUS_COLORS["TBC"]),
    ]
    legend_left   = left_origin
    legend_width  = SLIDE_W - LEFT_PAD - RIGHT_PAD
    slot_w        = legend_width // len(legend_items)
    legend_y      = LEGEND_TOP

    for i, (label, shp_type, color) in enumerate(legend_items):
        x = legend_left + i*slot_w + (slot_w - Inches(0.3))//2
        m = slide.shapes.add_shape(shp_type, x, legend_y, Inches(0.30), Inches(0.30))
        m.fill.solid()
        m.fill.fore_color.rgb = color
        m.line.fill.background()

        tb = slide.shapes.add_textbox(x + Inches(0.35), legend_y - Inches(0.02), slot_w - Inches(0.45), Inches(0.34))
        p = tb.text_frame.paragraphs[0]
        p.text = label
        p.font.size = Pt(15)
        p.alignment = PP_ALIGN.LEFT

    # ----- 3.2 Month header bar (Jan–Dec) -----
    # background bar across the grid
    header_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left_origin + SIDEBAR_W, top_origin,
        col_count * col_w, HEADER_H
    )
    header_bar.fill.solid(); header_bar.fill.fore_color.rgb = BLUE_HDR
    header_bar.line.fill.solid(); header_bar.line.fill.fore_color.rgb = WHITE; header_bar.line.width = Pt(0.0)

    # month cells + labels
    for m_idx, m_date in enumerate(pd.date_range(f"{year}-01-01", f"{year}-12-01", freq="MS")):
        cx = left_origin + SIDEBAR_W + m_idx*col_w
        cell = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, top_origin, col_w, HEADER_H)
        cell.fill.background()  # header is already blue bar; draw border only
        cell.line.fill.solid(); cell.line.fill.fore_color.rgb = WHITE; cell.line.width = Pt(0.95)

        tb = slide.shapes.add_textbox(cx, top_origin, col_w, HEADER_H)
        tf = tb.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = m_date.strftime("%b %y")
        p.font.bold = True; p.font.size = Pt(18)
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    # ----- 3.3 Sidebar table (Type, Workstream) -----
    tbl_rows = int(row_count + 1)  # +1 header
    tbl_cols = int(2)

    tbl_shape = slide.shapes.add_table(
        tbl_rows, tbl_cols,
        left_origin, top_origin,
        SIDEBAR_W, HEADER_H + row_count*row_h
    )
    tbl = tbl_shape.table
    tbl.columns[0].width = TYPE_COL_W
    tbl.columns[1].width = WORK_COL_W
    tbl.rows[0].height = HEADER_H
    for r in range(1, tbl_rows):
        tbl.rows[r].height = row_h

    # header cells
    for c, title in enumerate(("Type","Workstream")):
        cell = tbl.cell(0, c)
        cell.text = title
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True; p.font.size = Pt(18)
        p.alignment = PP_ALIGN.CENTER
        cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_HDR
        for border in (cell.border_left, cell.border_right, cell.border_top, cell.border_bottom):
            border.line_format.fill.solid()
            border.line_format.fill.fore_color.rgb = WHITE
            border.line_format.width = Pt(0.95)

    # body cells + alternating shading
    for r, grp in enumerate(groups, start=1):
        type_txt, work_txt = grp.split("\n")
        shade = COL_LIGHT if r % 2 == 1 else COL_DARK
        for c, txt in enumerate((type_txt, work_txt)):
            cell = tbl.cell(r, c)
            cell.text = txt
            cell.fill.solid(); cell.fill.fore_color.rgb = shade
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(15); p.alignment = PP_ALIGN.CENTER
            p.font.color.rgb = BLACK
            for border in (cell.border_left, cell.border_right, cell.border_top, cell.border_bottom):
                border.line_format.fill.solid()
                border.line_format.fill.fore_color.rgb = WHITE
                border.line_format.width = Pt(0.75)

    # ----- 3.4 Date grid (rectangles + borders) -----
    grid_top = top_origin + HEADER_H
    for i in range(col_count):
        for r in range(row_count):
            gx = left_origin + SIDEBAR_W + i*col_w
            gy = grid_top + r*row_h
            cell = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, gx, gy, col_w, row_h)
            cell.fill.solid()
            cell.fill.fore_color.rgb = COL_LIGHT if i % 2 == 0 else COL_DARK
            cell.line.fill.solid(); cell.line.fill.fore_color.rgb = WHITE; cell.line.width = Pt(0.95)

    # center navy guide line through each row
    for r in range(row_count):
        y_center = grid_top + r*row_h + row_h//2
        ln = slide.shapes.add_shape(MSO_CONNECTOR.STRAIGHT, left_origin + SIDEBAR_W, y_center, 0, 0)
        ln.end_x = left_origin + SIDEBAR_W + col_count*col_w
        ln.end_y = y_center
        ln.line.fill.solid(); ln.line.fill.fore_color.rgb = NAVY; ln.line.width = Pt(0.5)

    # ----- 3.5 Plot milestones (accurate within month) -----
    shape_map = {"Regular": MSO_SHAPE.OVAL, "Major": MSO_SHAPE.STAR_5_POINT}

    for _, row in subset_df.iterrows():
        dt = row["Milestone Date"]
        month_idx = dt.month - 1
        days_in_month = monthrange(dt.year, dt.month)[1]
        # fraction across month (0 at 1st, 1 at last day)
        day_frac = (dt.day - 1) / float(max(1, days_in_month - 1))

        # x/y for the marker
        x = left_origin + SIDEBAR_W + month_idx*col_w + int(day_frac * (col_w - Inches(0.01)))  # slight padding
        y_group = f'{row["Type"]}\n({row["Workstream"]})'
        yi = groups.index(y_group)
        y = grid_top + yi*row_h + row_h//2

        # choose size
        is_major = str(row.get("Milestone Type","Regular")).strip().lower() == "major"
        size = STAR_SIZE if is_major else CIRCLE_SIZE
        half = size // 2

        # center marker on (x,y)
        shp = slide.shapes.add_shape(shape_map["Major" if is_major else "Regular"], x - half, y - half, size, size)
        shp.fill.solid()
        shp.fill.fore_color.rgb = STATUS_COLORS.get(row["Milestone Status"], RGBColor(128,128,128))
        shp.line.color.rgb = BLACK

        # label just to the right & slightly above baseline
        lbl = slide.shapes.add_textbox(x + Inches(0.40), y - Inches(0.12), Inches(2.2), Inches(0.35))
        tf = lbl.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = str(row["Milestone Title"])
        p.font.size = Pt(12)
        p.font.color.rgb = BLACK

    # ----- 3.6 Current date (only if this slide is the current year) -----
    today = datetime.today().date()
    if today.year == year:
        dim = monthrange(today.year, today.month)[1]
        month_idx = today.month - 1
        day_frac = (today.day - 1) / float(max(1, dim - 1))
        xpos = left_origin + SIDEBAR_W + month_idx*col_w + int(day_frac * (col_w - Inches(0.01)))

        today_conn = slide.shapes.add_shape(MSO_CONNECTOR.STRAIGHT, xpos, top_origin + HEADER_H, 0, 0)
        today_conn.end_x = xpos
        today_conn.end_y = grid_top + row_count*row_h
        today_conn.line.fill.solid()
        today_conn.line.fill.fore_color.rgb = GREEN_DOT
        today_conn.line.width = Pt(2)
        today_conn.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

    # ----- 3.7 Title (optional page indicator) -----
    title = slide.shapes.add_textbox(LEFT_PAD, Inches(0.05), SLIDE_W - LEFT_PAD - RIGHT_PAD, Inches(0.5))
    p = title.text_frame.paragraphs[0]
    p.text = f"TM‑US Roadmap — {year}" + (f"  (Page {page_num}/{total_pages})" if total_pages > 1 else "")
    p.font.size = Pt(20); p.font.bold = True
    p.alignment = PP_ALIGN.LEFT


# -------------------------------------------------------------------
# 4) BUILD PRESENTATION ACROSS YEARS (and pages per year)
# -------------------------------------------------------------------

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

for year in sorted(df["Year"].unique()):
    df_year = df[df["Year"] == year].copy()

    # sorting already applied globally; re-derive groups in order
    df_year["Group"] = df_year["Type"].astype(str) + "\n(" + df_year["Workstream"].astype(str) + ")"

    # paginate by ROWS_PER_SLIDE
    groups_year = df_year[["Type","Workstream"]].drop_duplicates()
    # build group order for stable slicing
    groups_year["Group"] = groups_year["Type"] + "\n(" + groups_year["Workstream"] + ")"
    ordered_groups = groups_year["Group"].tolist()

    pages = [ordered_groups[i:i+ROWS_PER_SLIDE] for i in range(0, len(ordered_groups), ROWS_PER_SLIDE)]
    total_pages = max(1, len(pages))

    for idx, group_slice in enumerate(pages, start=1):
        # rows for this page
        sub = df_year[df_year["Type"].astype(str) + "\n(" + df_year["Workstream"].astype(str) + ")" \
                      .isin(group_slice)].copy()

        # and ensure milestone rows are ordered by the page's group order
        sub["Group"] = sub["Type"].astype(str) + "\n(" + sub["Workstream"].astype(str) + ")"
        sub["__gidx"] = sub["Group"].apply(lambda g: group_slice.index(g))
        sub = sub.sort_values(by=["__gidx", "Milestone Date"], kind="stable").drop(columns="__gidx")

        build_slide(prs, sub, year, idx, total_pages)

# -------------------------------------------------------------------
# 5) SAVE
# -------------------------------------------------------------------
prs.save(OUT_PPTX)
print("Saved:", OUT_PPTX)