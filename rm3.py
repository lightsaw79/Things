# --- imports
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

import pandas as pd
import re
from datetime import datetime, date
from calendar import monthrange

# ========== 1) INPUT ==========
XLSX_PATH = r"C:\Users\you\Desktop\Roadmap_Input.xlsx"   # <- change to your file
OUT_PPTX  = r"C:\Users\you\Desktop\Roadmap.pptx"

# Expected columns in Excel:
# Type | Workstream | Milestone Title | Milestone Date | Milestone Status | Milestone Type

# ========== 2) SIZING / LOOK ==========
SLIDE_W_IN = 20.0
SLIDE_H_IN = 9.0

TOP_PAD_IN     = 0.40  # gap above legend
LEGEND_H_IN    = 0.60
HEADER_H_IN    = 1.00  # month header height
LEFT_PAD_IN    = 0.25
RIGHT_PAD_IN   = 0.25
BOTTOM_PAD_IN  = 0.10

SIDEBAR_W_IN   = 4.00   # total width of the left sidebar (Type + Workstream)
TYPE_COL_W_IN  = 1.50   # inside the sidebar
WORK_COL_W_IN  = SIDEBAR_W_IN - TYPE_COL_W_IN

ROW_H_IN       = 0.60   # row height; more rows will still paginate (30/slide)
PAGE_ROWS      = 30

# Shapes
CIRCLE_SIZE_IN = 0.30   # regular milestone size
STAR_SIZE_IN   = 0.40   # major milestone size
LABEL_DX_IN    = 0.40   # label offset right
LABEL_DY_IN    = 0.10   # label offset up

# Colors
BLUE_HDR   = RGBColor(91,155,213)
MONTH_ODD  = RGBColor(224,242,255)
MONTH_EVEN = RGBColor(190,220,240)
WHITE      = RGBColor(255,255,255)
BLACK      = RGBColor(0,0,0)
NAVY       = RGBColor(0,0,128)
TODAY_GRN  = RGBColor(0,176,80)

STATUS_COLORS = {
    "On Track":  RGBColor(0,176,80),
    "At Risk":   RGBColor(255,192,0),
    "Off Track": RGBColor(255,0,0),
    "Complete":  RGBColor(0,112,192),
    "TBC":       RGBColor(191,191,191),
}

# ========== 3) READ & PREPARE DATA ==========
df = pd.read_excel(XLSX_PATH)
df["Milestone Date"] = pd.to_datetime(df["Milestone Date"], errors="coerce")
df["year"] = df["Milestone Date"].dt.year

# cleaning helpers (for consistent sorting)
def _clean(s):
    s = "" if pd.isna(s) else str(s)
    return s.strip().replace("\n"," ").casefold()

def _type_bucket(s):
    s = _clean(s)
    # first alnum run of the cleaned string
    m = re.match(r"([a-z0-9]+)", re.sub(r"[^a-z0-9]", "", s))
    return m.group(1) if m else s

# full sort for deterministic y-order
df["Type_key"]    = df["Type"].map(_clean).fillna("")
df["Type_bucket"] = df["Type"].map(_type_bucket).fillna("")
df["Work_key"]    = df["Workstream"].map(_clean).fillna("")
df = df.sort_values(by=["Type_bucket","Type_key","Work_key","Milestone Date"], kind="stable")

# y‑axis label text
df["Group"] = df["Type"].astype(str) + "\n(" + df["Workstream"].astype(str) + ")"

# ========== 4) PRESENTATION BOOTSTRAP ==========
prs = Presentation()
prs.slide_width  = Inches(SLIDE_W_IN)
prs.slide_height = Inches(SLIDE_H_IN)

# convenience floats → Inches at the last moment
def IN(x): return Inches(x)

# ========== helpers ==========
def add_full_width_legend(slide, left_in, top_in, width_in, height_in):
    """Draw legend spanning across the slide width above the table."""
    items = [
        ("Major Milestone", MSO_SHAPE.STAR_5_POINT,  STATUS_COLORS["On Track"]), # star colored green in your example
        ("On Track",        MSO_SHAPE.OVAL,         STATUS_COLORS["On Track"]),
        ("At Risk",         MSO_SHAPE.OVAL,         STATUS_COLORS["At Risk"]),
        ("Off Track",       MSO_SHAPE.OVAL,         STATUS_COLORS["Off Track"]),
        ("Complete",        MSO_SHAPE.OVAL,         STATUS_COLORS["Complete"]),
        ("TBC",             MSO_SHAPE.OVAL,         STATUS_COLORS["TBC"]),
    ]
    cols = len(items)
    slot_w_in = width_in / cols

    for i, (label, shp_kind, color) in enumerate(items):
        cx_in = left_in + i*slot_w_in + 0.30
        cy_in = top_in  + (height_in - 0.30)/2.0
        shp = slide.shapes.add_shape(shp_kind, IN(cx_in), IN(cy_in), IN(0.30), IN(0.30))
        shp.fill.solid(); shp.fill.fore_color.rgb = color
        shp.line.fill.background()

        tb = slide.shapes.add_textbox(IN(cx_in+0.35), IN(cy_in-0.02), IN(1.8), IN(0.30))
        p = tb.text_frame.paragraphs[0]
        p.text = label
        p.font.size = Pt(14)

def build_table_get_geometry(slide, groups, year):
    """
    Build one editable table (header row + data rows; 2 left columns + 12 months),
    color and border the cells, and return geometry numbers we need to position shapes.
    Returns:
      tbl, row_cells_top_in, month_left_in, month_w_in, first_row_top_in
    """
    total_width_in  = SLIDE_W_IN - LEFT_PAD_IN - RIGHT_PAD_IN
    total_height_in = SLIDE_H_IN - TOP_PAD_IN - LEGEND_H_IN - BOTTOM_PAD_IN
    top_origin_in   = TOP_PAD_IN + LEGEND_H_IN
    left_origin_in  = LEFT_PAD_IN

    # table shape
    row_count = len(groups)
    rows = 1 + row_count               # +1 header
    cols = 2 + 12                      # Type, Workstream, Jan..Dec

    tbl_shape = slide.shapes.add_table(
        rows, cols,
        IN(left_origin_in),
        IN(top_origin_in),
        IN(total_width_in),
        IN(total_height_in)
    )
    tbl = tbl_shape.table

    # column widths (keep months equal)
    tbl.columns[0].width = IN(TYPE_COL_W_IN)
    tbl.columns[1].width = IN(WORK_COL_W_IN)
    months_w_in = (total_width_in - SIDEBAR_W_IN) / 12.0
    for c in range(2, 14):
        tbl.columns[c].width = IN(months_w_in)

    # row heights
    tbl.rows[0].height = IN(HEADER_H_IN)
    for r in range(1, rows):
        tbl.rows[r].height = IN(ROW_H_IN)

    # header fills & text
    hdr_titles = ["Type","Workstream"] + [datetime(year, m, 1).strftime("%b %y") for m in range(1,13)]
    for c, title in enumerate(hdr_titles):
        cell = tbl.cell(0, c)
        cell.text = title
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(18)
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER
        cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_HDR
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    # body: text for Type/Workstream; alternating month fills; borders
    for r, grp in enumerate(groups, start=1):
        # left two columns (black text)
        t, w = grp.split("\n", 1)
        for c, value in enumerate([t, w]):
            cell = tbl.cell(r, c)
            cell.text = value
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(15)
            p.font.color.rgb = BLACK
            p.alignment = PP_ALIGN.CENTER
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            # left cells use a very light fill
            cell.fill.solid(); cell.fill.fore_color.rgb = MONTH_ODD

        # date cells: alternate by row; thin white borders
        for c in range(2, 14):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = MONTH_ODD if (r % 2 == 1) else MONTH_EVEN
            for b in (cell.border_left, cell.border_right, cell.border_top, cell.border_bottom):
                b.fill.solid()
                b.fill.fore_color.rgb = WHITE

    # compute month column left & width in *inches* for shape placement
    # (table gives us widths; we know table left; compute cumulative)
    month_left_in = left_origin_in + SIDEBAR_W_IN
    month_w_in    = months_w_in
    first_row_top_in = top_origin_in + HEADER_H_IN

    return tbl, first_row_top_in, month_left_in, month_w_in, top_origin_in, left_origin_in, total_width_in

def draw_row_center_lines(slide, row_count, left_in, right_in, first_row_top_in):
    """Thin navy line across the date area at the vertical center of each row."""
    for r in range(row_count):
        y_center_in = first_row_top_in + r*ROW_H_IN + ROW_H_IN/2.0
        ln = slide.shapes.add_shape(
            MSO_CONNECTOR.STRAIGHT,
            IN(left_in),  IN(y_center_in),
            IN(right_in - left_in), IN(0)
        )
        ln.line.fill.solid()
        ln.line.fill.fore_color.rgb = NAVY
        ln.line.width = Pt(0.5)

def draw_today_line_if_year(slide, year, month_left_in, month_w_in, first_row_top_in, row_count):
    """Dotted green vertical line at today's position within its month, only if year matches."""
    today = date.today()
    if today.year != year:
        return
    dim = monthrange(today.year, today.month)[1]
    day_frac = (today.day - 1) / (dim - 1)
    xpos_in = month_left_in + (today.month - 1 + day_frac)*month_w_in
    ln = slide.shapes.add_shape(
        MSO_CONNECTOR.STRAIGHT,
        IN(xpos_in), IN(first_row_top_in - ROW_H_IN),   # start a bit above rows (into header)
        IN(0),       IN(ROW_H_IN*row_count + ROW_H_IN)  # span header + rows
    )
    ln.line.fill.solid()
    ln.line.fill.fore_color.rgb = TODAY_GRN
    ln.line.width = Pt(2)
    ln.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

def plot_milestones(slide, df_year_page, groups, month_left_in, month_w_in, first_row_top_in):
    """Place stars/circles with accurate within-month x, and labels slightly above-right."""
    for _, row in df_year_page.iterrows():
        dt = row["Milestone Date"]
        if pd.isna(dt): 
            continue

        month_idx = dt.month - 1
        dim = monthrange(dt.year, dt.month)[1]
        day_frac = (dt.day - 1) / (dim - 1)                # 0..1 across the month
        x_in = month_left_in + (month_idx + day_frac)*month_w_in

        # y position
        g_idx = groups.index(row["Group"])                 # row 0-based within the page
        y_in = first_row_top_in + g_idx*ROW_H_IN + ROW_H_IN/2.0

        # choose shape + size
        is_major = str(row.get("Milestone Type","")).strip().lower() == "major"
        size_in  = STAR_SIZE_IN if is_major else CIRCLE_SIZE_IN
        half     = size_in/2.0

        shp_kind = MSO_SHAPE.STAR_5_POINT if is_major else MSO_SHAPE.OVAL
        shp = slide.shapes.add_shape(shp_kind, IN(x_in - half), IN(y_in - half), IN(size_in), IN(size_in))
        shp.fill.solid(); shp.fill.fore_color.rgb = STATUS_COLORS.get(row.get("Milestone Status","On Track"), BLACK)
        shp.line.color.rgb = BLACK

        # label (slightly above and right)
        tb = slide.shapes.add_textbox(IN(x_in + LABEL_DX_IN), IN(y_in - LABEL_DY_IN), IN(2.0), IN(0.35))
        p  = tb.text_frame.paragraphs[0]
        p.text = str(row.get("Milestone Title","")).strip()
        p.font.size = Pt(12)
        p.font.color.rgb = BLACK

# ========== 5) BUILD SLIDES (one per year; paginate 30 rows/slide) ==========
for year in sorted(df["year"].dropna().unique()):
    df_year = df.loc[df["year"] == year].copy()

    # Build the display order of groups for this year (unique list in sorted order)
    group_list = (
        df_year[["Type_bucket","Type_key","Work_key","Group"]]
        .drop_duplicates()
        .sort_values(by=["Type_bucket","Type_key","Work_key"], kind="stable")
        .apply(lambda s: f"{df_year.loc[s.name,'Group']}", axis=1)  # keep the original label
    )

    # If the previous line looks odd in your pandas version, this simpler one works too:
    group_list = (
        df_year[["Type_bucket","Type_key","Work_key","Group"]]
        .drop_duplicates()
        .sort_values(by=["Type_bucket","Type_key","Work_key"], kind="stable")["Group"]
        .tolist()
    )

    total_rows = len(group_list)
    pages = [group_list[i:i+PAGE_ROWS] for i in range(0, total_rows, PAGE_ROWS)] or [[]]

    for page_no, groups in enumerate(pages, start=1):
        # New slide
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Legend
        add_full_width_legend(
            slide,
            left_in=LEFT_PAD_IN,
            top_in=TOP_PAD_IN,
            width_in=SLIDE_W_IN - LEFT_PAD_IN - RIGHT_PAD_IN,
            height_in=LEGEND_H_IN
        )

        # Table + geometry
        tbl, first_row_top_in, month_left_in, month_w_in, top_origin_in, left_origin_in, total_w_in = \
            build_table_get_geometry(slide, groups, year)

        # Navy row-center lines inside dates area (left/right across the 12 month area)
        dates_left_in  = month_left_in
        dates_right_in = month_left_in + 12*month_w_in
        draw_row_center_lines(slide, len(groups), dates_left_in, dates_right_in, first_row_top_in)

        # Slice data for this page (keep milestone order stable within page groups)
        sub = df_year[df_year["Group"].isin(groups)].copy()
        # Ensure within-page y order matches 'groups'
        order_index = {g:i for i,g in enumerate(groups)}
        sub["_gidx"] = sub["Group"].map(order_index)
        sub = sub.sort_values(by=["_gidx","Milestone Date"], kind="stable").drop(columns="_gidx")

        # Plot milestones
        plot_milestones(slide, sub, groups, month_left_in, month_w_in, first_row_top_in)

        # Today dotted vertical (only for current year)
        draw_today_line_if_year(slide, year, month_left_in, month_w_in, first_row_top_in, len(groups))

        # Optional slide title
        title_tb = slide.shapes.add_textbox(IN(LEFT_PAD_IN), IN(0.05), IN(8.0), IN(0.5))
        pt = title_tb.text_frame.paragraphs[0]
        pt.text = f"TM‑US Roadmap — {year}  (Page {page_no}/{len(pages)})"
        pt.font.size = Pt(22)
        pt.font.bold = True
        pt.font.color.rgb = BLACK

# ========== 6) SAVE ==========
prs.save(OUT_PPTX)
print("Saved:", OUT_PPTX)