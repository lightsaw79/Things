from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

import pandas as pd
from datetime import datetime, date
from calendar import monthrange
import math
import re

# ---------- FILES ----------
INPUT_XLSX = r"C:\path\to\Roadmap_Input.xlsx"
OUT_PPTX   = r"C:\path\to\Roadmap.pptx"

# ---------- READ & PREPARE DATA ----------
df = pd.read_excel(INPUT_XLSX)
# Expected columns: Type, Workstream, Milestone Title, Milestone Date, Milestone Status, Milestone Type
df['Milestone Date'] = pd.to_datetime(df['Milestone Date'])
df['year'] = df['Milestone Date'].dt.year

# Sorting helpers: bucket on leading word/letters+digits of Type
def clean(s): 
    return str(s).strip().replace("\n", " ").casefold()

def type_bucket(s):
    s = clean(s)
    m = re.match(r"([a-z0-9]+)", re.sub(r"[^a-z0-9]", "", s))
    return m.group(1) if m else s

df['Type_key'] = df['Type'].map(clean)
df['Type_bucket'] = df['Type'].map(type_bucket)
df['Work_key'] = df['Workstream'].map(clean)

df = df.sort_values(
    by=['Type_bucket', 'Type_key', 'Work_key', 'Milestone Date'],
    kind='stable'
)

# Group label shown on y-axis
df['Group'] = df['Type'].astype(str) + "\n(" + df['Workstream'].astype(str) + ")"

# ---------- CONSTANTS (layout) ----------
SLIDE_W = Inches(20)   # keep your wide slide
SLIDE_H = Inches(9)

TOP_PAD     = Inches(0.4)   # gap above legend
LEGEND_H    = Inches(0.6)
HEADER_H    = Inches(1.0)   # month header row height
BOTTOM_PAD  = Inches(0.09)  # keep small gap at bottom

LEFT_PAD    = Inches(0.25)  # inner left padding inside the table
RIGHT_PAD   = Inches(0.25)

SIDEBAR_W   = Inches(4.0)   # total width for Type + Workstream
TYPE_COL_W  = Inches(1.5)
WORK_COL_W  = SIDEBAR_W - TYPE_COL_W

MONTHS = 12
COL_COUNT = 2 + MONTHS      # Type, Workstream + months

# Navy row center line
ROW_LINE_COLOR = RGBColor(0, 0, 128)
ROW_LINE_W_PT  = Pt(0.5)

# Month header color
BLUE_HDR = RGBColor(91, 155, 213)

# Alternating row fill colors
ROW_LIGHT = RGBColor(224, 242, 255)
ROW_DARK  = RGBColor(190, 220, 240)

# Status colors
STATUS_COLORS = {
    "On Track": RGBColor(0,176,80),
    "At Risk":  RGBColor(255,192,0),
    "Off Track":RGBColor(255,0,0),
    "Complete": RGBColor(0,112,192),
    "TBC":      RGBColor(191,191,191),
}

# Milestone sizes
CIRCLE_SIZE = Inches(0.30)
STAR_SIZE   = Inches(0.40)   # change here to resize Major star

# Label offsets
LABEL_DX = Inches(0.40)
LABEL_DY = Inches(0.10)

# Rows per page
ROWS_PER_PAGE = 30

# ---------- HELPERS ----------

def add_legend(slide):
    """Draw a full-width legend row at the very top."""
    legend_items = [
        ("Major Milestone", MSO_SHAPE.STAR_5_POINT, STATUS_COLORS["On Track"]), # star colored green marker
        ("On Track",  MSO_SHAPE.OVAL, STATUS_COLORS["On Track"]),
        ("At Risk",   MSO_SHAPE.OVAL, STATUS_COLORS["At Risk"]),
        ("Off Track", MSO_SHAPE.OVAL, STATUS_COLORS["Off Track"]),
        ("Complete",  MSO_SHAPE.OVAL, STATUS_COLORS["Complete"]),
        ("TBC",       MSO_SHAPE.OVAL, STATUS_COLORS["TBC"]),
    ]
    left = int(LEFT_PAD)
    top  = int(TOP_PAD)
    total = SLIDE_W - LEFT_PAD - RIGHT_PAD
    slot = total / len(legend_items)

    for i,(label, shp_type, col) in enumerate(legend_items):
        x = int(LEFT_PAD + i*slot + Inches(0.1))
        y = top
        size = int(Inches(0.3))
        shp = slide.shapes.add_shape(shp_type, x, y, size, size)
        shp.fill.solid()
        shp.fill.fore_color.rgb = col
        shp.line.fill.background()

        tb = slide.shapes.add_textbox(int(x + size + Inches(0.15)),
                                      y - int(Inches(0.03)),
                                      int(Inches(2.5)),
                                      size)
        p = tb.text_frame.paragraphs[0]
        p.text = label
        p.font.size = Pt(15)

def month_label(dt):
    return dt.strftime("%b %y")

def build_full_editable_table(slide, year, groups, total_width, total_height):
    """
    Build a true PowerPoint table covering the whole grid area (left sidebar + 12 months).
    Returns geometry needed to place shapes independently of the table.
    """
    row_count = len(groups)

    # Table anchor & size
    left   = int(LEFT_PAD)
    top    = int(TOP_PAD + LEGEND_H)
    width  = int(total_width)
    height = int(total_height)

    # Create table
    tbl_rows = int(row_count + 1)   # +1 header
    tbl_cols = int(COL_COUNT)       # 2 + 12

    tbl_shape = slide.shapes.add_table(tbl_rows, tbl_cols, left, top, width, height)
    tbl = tbl_shape.table

    # Column widths
    # First two are fixed; remaining equally share the rest
    dates_width_total = total_width - SIDEBAR_W - LEFT_PAD - RIGHT_PAD
    month_w = dates_width_total / MONTHS

    tbl.columns[0].width = int(TYPE_COL_W)
    tbl.columns[1].width = int(WORK_COL_W)
    for i in range(MONTHS):
        tbl.columns[2+i].width = int(month_w)

    # Row heights
    tbl.rows[0].height = int(HEADER_H)
    body_h = total_height - HEADER_H
    row_h  = body_h / max(row_count,1)
    for r in range(1, tbl_rows):
        tbl.rows[r].height = int(row_h)

    # Header cells (month names)
    # paint header background
    for c in range(tbl_cols):
        cell = tbl.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE_HDR
        tf = cell.text_frame
        p = tf.paragraphs[0]
        if c == 0:
            p.text = "Type"
        elif c == 1:
            p.text = "Workstream"
        else:
            month_dt = date(year, c-1, 1)  # c-2 + 1 = c-1
            p.text = month_label(pd.Timestamp(year=year, month=c-1, day=1))
        p.font.color.rgb = RGBColor(255,255,255)
        p.font.size = Pt(18)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p.alignment = PP_ALIGN.CENTER

    # Sidebar cells (Type / Workstream) + alternating row fill
    for r, grp in enumerate(groups, start=1):
        type_txt, work_txt = grp.split("\n")
        # alternating fill for the whole row of date area (we’ll color Type/Work with same tone)
        row_color = ROW_LIGHT if (r % 2)==1 else ROW_DARK
        for c in range(tbl_cols):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_color

        # Type text
        tcell = tbl.cell(r, 0)
        tcell.text_frame.paragraphs[0].text = type_txt
        tcell.text_frame.paragraphs[0].font.size = Pt(15)
        tcell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        tcell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        # Workstream text
        wcell = tbl.cell(r, 1)
        wcell.text_frame.paragraphs[0].text = work_txt.strip("()")
        wcell.text_frame.paragraphs[0].font.size = Pt(15)
        wcell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        wcell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # We cannot set table cell borders with high-level API; draw thin white lines across columns/rows:
    # vertical grid
    x_cursor = left + int(TYPE_COL_W + WORK_COL_W)
    for m in range(MONTHS+1):
        # vertical gridline from below header to bottom
        x = int(left + TYPE_COL_W + WORK_COL_W + m*month_w)
        line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            x, int(top + HEADER_H),
            x, int(top + HEADER_H + row_count*row_h)
        )
        line.line.fill.solid()
        line.line.fill.fore_color.rgb = RGBColor(255,255,255)
        line.line.width = Pt(0.95)

    # horizontal grid (between body rows)
    for r in range(row_count+1):
        y = int(top + HEADER_H + r*row_h)
        line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            int(left + TYPE_COL_W + WORK_COL_W),
            y,
            int(left + TYPE_COL_W + WORK_COL_W + MONTHS*month_w),
            y
        )
        line.line.fill.solid()
        line.line.fill.fore_color.rgb = RGBColor(255,255,255)
        line.line.width = Pt(0.95)

    # Navy center lines through each body row (dates area only)
    for r in range(row_count):
        y_center = int(top + HEADER_H + r*row_h + row_h/2)
        ln = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            int(left + TYPE_COL_W + WORK_COL_W),
            y_center,
            int(left + TYPE_COL_W + WORK_COL_W + MONTHS*month_w),
            y_center
        )
        ln.line.fill.solid()
        ln.line.fill.fore_color.rgb = ROW_LINE_COLOR
        ln.line.width = ROW_LINE_W_PT

    # return geometry used for independent shapes
    dates_left  = int(left + TYPE_COL_W + WORK_COL_W)
    dates_right = int(dates_left + MONTHS*month_w)
    return {
        "top": int(top),
        "left": int(left),
        "row_h": float(row_h),
        "dates_top": int(top + HEADER_H),
        "dates_left": dates_left,
        "month_w": float(month_w),
        "row_count": row_count,
        "dates_bottom": int(top + HEADER_H + row_count*row_h),
    }

def add_today_line(slide, geom, year):
    """Dotted green vertical line at actual day position, only on matching year slide."""
    today = date.today()
    if today.year != year:
        return
    days_in_month = monthrange(today.year, today.month)[1]
    # fraction across the month (0…1)
    day_frac = (today.day - 1) / (days_in_month - 1) if days_in_month > 1 else 0
    month_idx = today.month - 1
    x = int(geom["dates_left"] + (month_idx + day_frac) * geom["month_w"])
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        x, int(geom["dates_top"]),
        x, int(geom["dates_bottom"])
    )
    conn.line.fill.solid()
    conn.line.fill.fore_color.rgb = RGBColor(0,176,80)
    conn.line.width = Pt(2)
    conn.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

def place_milestones(slide, geom, df_subset, groups):
    """Place independent shapes (circles/stars) and labels at accurate day positions."""
    for _, row in df_subset.iterrows():
        dt = pd.Timestamp(row['Milestone Date'])
        month_idx = dt.month - 1
        dim = monthrange(dt.year, dt.month)[1]
        day_frac = (dt.day - 1) / (dim - 1) if dim > 1 else 0.0

        # X position (center of the marker)
        x_center = geom["dates_left"] + (month_idx + day_frac) * geom["month_w"]
        # Y center = row center
        y_index = groups.index(row['Group'])
        y_center = geom["dates_top"] + y_index * geom["row_h"] + geom["row_h"]/2

        # shape choice & size
        is_major = str(row.get('Milestone Type','')).strip().lower() == "major"
        size = STAR_SIZE if is_major else CIRCLE_SIZE
        half = size/2.0

        shp = slide.shapes.add_shape(
            MSO_SHAPE.STAR_5_POINT if is_major else MSO_SHAPE.OVAL,
            int(x_center - half),
            int(y_center - half),
            int(size),
            int(size)
        )
        # color
        col = STATUS_COLORS.get(str(row.get('Milestone Status','')).strip(), RGBColor(128,128,128))
        shp.fill.solid(); shp.fill.fore_color.rgb = col
        shp.line.color.rgb = RGBColor(0,0,0)

        # label (slightly right & above)
        lbl = slide.shapes.add_textbox(
            int(x_center + LABEL_DX),
            int(y_center - LABEL_DY),
            int(Inches(2.0)), int(Inches(0.35))
        )
        p = lbl.text_frame.paragraphs[0]
        p.text = str(row['Milestone Title'])
        p.font.size = Pt(12)

# ---------- BUILD SLIDES ----------
prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

for year in sorted(df['year'].unique()):
    df_year = df[df['year']==year].copy()

    # Build the display order of groups (Type/Workstream) for this year (unique, in sorted order)
    groups = (df_year[['Type_key','Type','Work_key','Workstream']]
              .drop_duplicates()
              .sort_values(by=['Type_bucket','Type_key','Work_key'], key=lambda s: s)
              .apply(lambda r: f"{r['Type']}\n({r['Workstream']})", axis=1)
              .tolist()
             )

    # paginate 30 rows per slide
    total_rows = len(groups)
    total_pages = math.ceil(max(total_rows,1)/ROWS_PER_PAGE)

    for page_idx in range(total_pages):
        start = page_idx*ROWS_PER_PAGE
        end   = min((page_idx+1)*ROWS_PER_PAGE, total_rows)
        groups_page = groups[start:end]

        # subset rows that belong to these groups, keep the same order
        order_index = {g:i for i,g in enumerate(groups_page)}
        sub = df_year[df_year['Group'].isin(groups_page)].copy()
        sub['_gid'] = sub['Group'].map(order_index)
        sub = sub.sort_values(by=['_gid','Milestone Date'], kind='stable').drop(columns=['_gid'])

        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # legend
        add_legend(slide)

        # available grid area (width/height) under legend
        grid_top = TOP_PAD + LEGEND_H
        total_height = SLIDE_H - grid_top - BOTTOM_PAD
        total_width  = SLIDE_W - LEFT_PAD - RIGHT_PAD

        # build editable table and compute geometry
        geom = build_full_editable_table(slide, year, groups_page, total_width, total_height)

        # title (top-left, above legend)
        title_tb = slide.shapes.add_textbox(int(LEFT_PAD), int(Inches(0.02)),
                                            int(Inches(8)), int(Inches(0.5)))
        tp = title_tb.text_frame.paragraphs[0]
        tp.text = f"TM-US Roadmap — {year}  (Rows {start+1}-{end} of {total_rows})"
        tp.font.size = Pt(22)

        # milestones
        place_milestones(slide, geom, sub, groups_page)

        # today line
        add_today_line(slide, geom, year)

# SAVE
prs.save(OUT_PPTX)
print("Saved:", OUT_PPTX)