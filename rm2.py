# --- Roadmap → editable PPT (one slide per year) -----------------------------
# What you get:
#   • Legend across the top
#   • Header row with months for that year
#   • Left sidebar with Type + Workstream, sorted A→Z
#   • Grid (12 months × N groups), thin borders
#   • Milestones: circle (Regular), star (Major), placed by day-of-month
#   • “Today” dotted green line on the current year slide
#   • Major star is larger than circles (resize values at the knobs below)

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

import pandas as pd
from calendar import monthrange
from datetime import datetime

# 0) ----------- INPUT EXCEL ---------------------------------------------------
# Columns expected (case-insensitive): Type, Workstream, Milestone Title,
# Milestone Date, Milestone Status, Milestone Type ("Regular"/"Major")
EXCEL_PATH = r"Roadmap_Input_Sheet.xlsx"

df = pd.read_excel(EXCEL_PATH)
df.rename(columns={c: c.strip() for c in df.columns}, inplace=True)
df["Milestone Date"] = pd.to_datetime(df["Milestone Date"])
df["Year"] = df["Milestone Date"].dt.year      # used to split slides by year

# 1) ----------- PRESENTATION CANVAS ------------------------------------------
prs = Presentation()
prs.slide_width  = Inches(20)  # ultra-wide
prs.slide_height = Inches(9)

# 2) ----------- VISUAL CONSTANTS (easy knobs) --------------------------------
# Sidebar widths
SIDEBAR_W   = Inches(4.0)     # Type+Workstream block total width
TYPE_COL_W  = Inches(1.5)
WORK_COL_W  = SIDEBAR_W - TYPE_COL_W

HEADER_H    = Inches(1.0)     # month header height
LEGEND_H    = Inches(0.6)
LEGEND_TOP  = Inches(0.2)
LEFT_PAD    = Inches(0.25)    # small left padding before grid
RIGHT_PAD   = Inches(0.50)    # right breathing room

TOP_MARGIN  = LEGEND_TOP + LEGEND_H + Inches(0.2)
BOTTOM_MARGIN = Inches(0.6)   # increase if the grid ever touches slide bottom

COL_COUNT   = 12              # 12 months

# Milestone sizes (change these to resize markers)
CIRCLE_SIZE = Inches(0.30)    # Regular
STAR_SIZE   = Inches(0.45)    # Major ⭐ (bigger than circles)

# Colors (RGB)
BLUE_HDR    = RGBColor(91,155,213)
ROW_EVEN    = RGBColor(242,242,242)
ROW_ODD     = RGBColor(190,220,240)
GRID_STROKE = RGBColor(255,255,255)
TEXT_WHITE  = RGBColor(255,255,255)
TEXT_BLACK  = RGBColor(0,0,0)
LINE_GREY   = RGBColor(128,128,128)
TODAY_GREEN = RGBColor(0,176,80)

# Status → color
STATUS_COLORS = {
    "On Track"  : RGBColor(0,176,80),
    "At Risk"   : RGBColor(255,192,0),
    "Off Track" : RGBColor(255,0,0),
    "Complete"  : RGBColor(0,112,192),
    "TBC"       : RGBColor(191,191,191),
}

# Milestone Type → shape
SHAPE_MAP = {"Regular": MSO_SHAPE.OVAL, "Major": MSO_SHAPE.STAR_5_POINT}


# 3) ----------- LOOP: ONE SLIDE PER YEAR -------------------------------------
for year in sorted(df["Year"].unique()):
    df_year = df[df["Year"] == year].copy()

    # 3a) SORT order for Y-axis: Type (A→Z) then Workstream (A→Z)
    df_year["Type_key"] = (
        df_year["Type"].fillna("").str.replace(r"\s+", " ", regex=True).str.strip().str.casefold()
    )
    df_year["Work_key"] = (
        df_year["Workstream"].fillna("").str.replace(r"\s+", " ", regex=True).str.strip().str.casefold()
    )
    # (optional bucket to keep close variants together; remove if not wanted)
    df_year["Type_bucket"] = (
        df_year["Type_key"].str.extract(r"^([a-z0-9]+)", expand=False).fillna(df_year["Type_key"])
    )
    df_year = df_year.sort_values(
        ["Type_bucket", "Type_key", "Work_key", "Milestone Date"],
        kind="stable"
    )
    df_year["Group"] = df_year["Type"] + "\n(" + df_year["Workstream"] + ")"
    groups = list(dict.fromkeys(df_year["Group"]))     # keep first occurrence order
    row_count = len(groups)

    # 3b) Geometry based on row count
    chart_w = prs.slide_width - SIDEBAR_W - RIGHT_PAD - LEFT_PAD
    chart_h = prs.slide_height - TOP_MARGIN - BOTTOM_MARGIN
    col_w   = chart_w / COL_COUNT
    row_h   = chart_h / max(row_count, 1)   # avoid zero-divide

    left_origin = SIDEBAR_W + LEFT_PAD
    top_origin  = TOP_MARGIN

    # Create slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 3c) Legend (spans the full width above the table)
    legend_items = [
        ("Major Milestone", MSO_SHAPE.STAR_5_POINT, STATUS_COLORS["On Track"]),  # star shown green
        ("On Track",        MSO_SHAPE.OVAL,        STATUS_COLORS["On Track"]),
        ("At Risk",         MSO_SHAPE.OVAL,        STATUS_COLORS["At Risk"]),
        ("Off Track",       MSO_SHAPE.OVAL,        STATUS_COLORS["Off Track"]),
        ("Complete",        MSO_SHAPE.OVAL,        STATUS_COLORS["Complete"]),
        ("TBC",             MSO_SHAPE.OVAL,        STATUS_COLORS["TBC"]),
    ]
    slots = len(legend_items)
    legend_left = Inches(0.5)
    legend_right = prs.slide_width - Inches(0.5)
    legend_slot_w = (legend_right - legend_left) / slots

    for i, (lbl, shp_type, col) in enumerate(legend_items):
        lx = legend_left + i * legend_slot_w
        # shape (slightly different size for star vs circle to preview)
        sz = Inches(0.35) if shp_type == MSO_SHAPE.STAR_5_POINT else Inches(0.28)
        m = slide.shapes.add_shape(shp_type, lx + Inches(0.1), LEGEND_TOP, sz, sz)
        m.fill.solid(); m.fill.fore_color.rgb = col
        m.line.fill.background()
        # label
        tb = slide.shapes.add_textbox(lx + Inches(0.5), LEGEND_TOP, legend_slot_w - Inches(0.6), Inches(0.35))
        p = tb.text_frame.paragraphs[0]
        p.text = lbl
        p.font.size = Pt(12)

    # 3d) Month header row (Jan→Dec of THIS year)
    for i, m in enumerate(pd.date_range(f"{year}-01-01", f"{year}-12-01", freq="MS")):
        cell = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_origin + i*col_w, top_origin, col_w, HEADER_H)
        cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_HDR
        cell.line.fill.solid(); cell.line.fill.fore_color.rgb = GRID_STROKE; cell.line.width = Pt(0.95)
        tf = cell.text_frame; tf.text = m.strftime("%b %y")
        p = tf.paragraphs[0]; p.font.bold = True; p.font.color.rgb = TEXT_WHITE; p.font.size = Pt(12); p.alignment = PP_ALIGN.CENTER
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    # 3e) Sidebar headers (“Type” / “Workstream”)
    th1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT_PAD, top_origin, TYPE_COL_W, HEADER_H)
    th1.fill.solid(); th1.fill.fore_color.rgb = BLUE_HDR
    th1.line.fill.solid(); th1.line.fill.fore_color.rgb = GRID_STROKE; th1.line.width = Pt(0.95)
    th1.text_frame.text = "Type"; p = th1.text_frame.paragraphs[0]
    p.font.bold = True; p.font.color.rgb = TEXT_WHITE; p.font.size = Pt(12); p.alignment = PP_ALIGN.CENTER
    th1.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    th2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT_PAD + TYPE_COL_W, top_origin, WORK_COL_W, HEADER_H)
    th2.fill.solid(); th2.fill.fore_color.rgb = BLUE_HDR
    th2.line.fill.solid(); th2.line.fill.fore_color.rgb = GRID_STROKE; th2.line.width = Pt(0.95)
    th2.text_frame.text = "Workstream"; p = th2.text_frame.paragraphs[0]
    p.font.bold = True; p.font.color.rgb = TEXT_WHITE; p.font.size = Pt(12); p.alignment = PP_ALIGN.CENTER
    th2.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # 3f) Sidebar rows (sorted groups), alternating fill
    for r, grp in enumerate(groups):
        y = top_origin + HEADER_H + r * row_h
        bg = ROW_EVEN if r % 2 == 0 else ROW_ODD
        type_txt, work_txt = grp.split("\n")

        # Type cell
        c1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT_PAD, y, TYPE_COL_W, row_h)
        c1.fill.solid(); c1.fill.fore_color.rgb = bg
        c1.line.fill.solid(); c1.line.fill.fore_color.rgb = GRID_STROKE; c1.line.width = Pt(0.5)
        tf = c1.text_frame; tf.text = type_txt; p = tf.paragraphs[0]
        p.font.size = Pt(10); p.font.color.rgb = TEXT_BLACK; p.alignment = PP_ALIGN.CENTER
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Workstream cell
        c2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT_PAD + TYPE_COL_W, y, WORK_COL_W, row_h)
        c2.fill.solid(); c2.fill.fore_color.rgb = bg
        c2.line.fill.solid(); c2.line.fill.fore_color.rgb = GRID_STROKE; c2.line.width = Pt(0.5)
        tf = c2.text_frame; tf.text = work_txt.strip("()"); p = tf.paragraphs[0]
        p.font.size = Pt(10); p.font.color.rgb = TEXT_BLACK; p.alignment = PP_ALIGN.CENTER
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    # 3g) Grid cells (white fill + thin white borders) and row separators
    for i in range(COL_COUNT):
        for r in range(row_count):
            x = left_origin + i * col_w
            y = top_origin + HEADER_H + r * row_h
            cell = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, col_w, row_h)
            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255,255,255)
            cell.line.fill.solid(); cell.line.fill.fore_color.rgb = GRID_STROKE; cell.line.width = Pt(0.5)

    # optional: thin horizontal lines across each row (helps alignment)
    for r in range(row_count + 1):
        y = top_origin + HEADER_H + r * row_h
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, left_origin, y, left_origin + chart_w, y)
        ln.line.fill.solid(); ln.line.fill.fore_color.rgb = LINE_GREY; ln.line.width = Pt(0.5)

    # 3h) Plot milestones (day-fraction inside month) + bigger star for Major
    for _, row in df_year.iterrows():
        dt = row["Milestone Date"]
        month_idx = dt.month - 1
        days_in = monthrange(dt.year, dt.month)[1]
        day_frac = (dt.day - 1) / float(days_in - 1) if days_in > 1 else 0.0

        # choose size by type; keep shape centered by subtracting half
        size = STAR_SIZE if str(row.get("Milestone Type", "Regular")).strip().lower() == "major" else CIRCLE_SIZE
        half = size / 2.0

        x = left_origin + (month_idx + day_frac) * col_w - half
        yi = groups.index(row["Group"])
        y = top_origin + HEADER_H + yi * row_h + row_h / 2.0 - half

        shp = slide.shapes.add_shape(SHAPE_MAP[row["Milestone Type"]], x, y, size, size)
        shp.fill.solid()
        shp.fill.fore_color.rgb = STATUS_COLORS.get(row["Milestone Status"], RGBColor(128,128,128))
        shp.line.fill.background()

        # label just right of the marker, slightly above midline
        lbl = slide.shapes.add_textbox(x + half + Inches(0.25), y - Inches(0.05), Inches(3), Inches(0.35))
        tf = lbl.text_frame
        tf.text = str(row["Milestone Title"])
        tf.paragraphs[0].font.size = Pt(10)
        tf.paragraphs[0].font.color.rgb = TEXT_BLACK

    # 3i) “Today” dotted green line (only on the current year slide)
    today = datetime.today().date()
    if today.year == year:
        dim = monthrange(today.year, today.month)[1]
        tfrac = (today.day - 1) / float(dim - 1) if dim > 1 else 0.0
        xpos = left_origin + (today.month - 1 + tfrac) * col_w
        tln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, xpos, top_origin + HEADER_H, xpos, top_origin + HEADER_H + chart_h)
        tln.line.fill.solid(); tln.line.fill.fore_color.rgb = TODAY_GREEN
        tln.line.width = Pt(1.0); tln.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

# 4) ----------- SAVE ----------------------------------------------------------
OUTPUT_PPTX = r"Roadmap_by_Year.pptx"
prs.save(OUTPUT_PPTX)
print(f"Saved → {OUTPUT_PPTX}")