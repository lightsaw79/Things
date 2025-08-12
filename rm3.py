# --- Roadmap to PPT (rectangles layout; accurate month-day placement) ---

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

import pandas as pd
from datetime import datetime
from calendar import monthrange

# ========= 1) INPUT =========
EXCEL_PATH = r"C:\Users\you\Desktop\Python Roadmap\IFC-Roadmap_Input_DRAFT_Sheet.xlsx"
SHEET_NAME = 0   # or "Sheet1"

df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
df["Milestone Date"] = pd.to_datetime(df["Milestone Date"])
df["Year"] = df["Milestone Date"].dt.year
df["Group"] = df["Type"] + "\n(" + df["Workstream"] + ")"

# ========= 2) CONSTANTS (keep your sizing/spacing pattern) =========
prs = Presentation()
prs.slide_width  = Inches(20)  # extra wide
prs.slide_height = Inches(9)

# Colors
BLUE_HDR      = RGBColor( 91,155,213)   # top month header & sidebar header
TEXT_WHITE    = RGBColor(255,255,255)
TEXT_BLACK    = RGBColor(  0,  0,  0)
ROW_EVEN      = RGBColor(224,242,255)   # light blue
ROW_ODD       = RGBColor(190,220,240)   # medium blue
CELL_BORDER   = RGBColor(255,255,255)   # thin white borders
ROW_CENTER    = RGBColor(  0,  0,128)   # navy center line
TODAY_GREEN   = RGBColor(  0,176, 80)

# Status → colors (keep your palette)
status_colors = {
    "On Track":   RGBColor(144,238,144),  # light green
    "At Risk":    RGBColor(255,192,  0),  # yellow
    "Off Track":  RGBColor(255,  0,  0),  # red
    "Complete":   RGBColor(  0,112,192),  # blue
    "TBC":        RGBColor(191,191,191),  # light gray
}

# Legend items (star for Major Milestone + status circles)
legend_items = [
    ("Major Milestone", MSO_SHAPE.STAR_5_POINT, RGBColor(192,164, 72)),  # dark golden
    ("On Track",        MSO_SHAPE.OVAL,         status_colors["On Track"]),
    ("At Risk",         MSO_SHAPE.OVAL,         status_colors["At Risk"]),
    ("Off Track",       MSO_SHAPE.OVAL,         status_colors["Off Track"]),
    ("Complete",        MSO_SHAPE.OVAL,         status_colors["Complete"]),
    ("TBC",             MSO_SHAPE.OVAL,         status_colors["TBC"]),
]

# Geometry (same spirit as your current file)
sidebar_w     = Inches(4)       # Type + Workstream total width
type_col_w    = Inches(1.5)
work_col_w    = sidebar_w - type_col_w
header_h      = Inches(1.0)
legend_height = Inches(0.6)
legend_top    = Inches(0.2)

# Margins inside slide
LEFT_PAD      = Inches(0.25)
RIGHT_PAD     = Inches(0.25)
top_margin    = legend_top + legend_height + Inches(0.2)
bottom_margin = Inches(0.5)    # ensures table ends above slide bottom

# Symbol sizes
CIRCLE_SIZE   = Inches(0.30)
STAR_SIZE     = Inches(0.40)

# Shape map
SHAPE_MAP = {"Regular": MSO_SHAPE.OVAL, "Major": MSO_SHAPE.STAR_5_POINT}

# ========= 3) PER-YEAR SLIDES =========
for year in sorted(df["Year"].unique()):
    df_year = df[df["Year"] == year].copy()

    # --- Sorting: Type (A→Z), then Workstream (A→Z) ---
    df_year["Type_key"] = (
        df_year["Type"].fillna("")
        .str.replace(r"\s+", " ", regex=True).str.strip().str.casefold()
    )
    df_year["Work_key"] = (
        df_year["Workstream"].fillna("")
        .str.replace(r"\s+", " ", regex=True).str.strip().str.casefold()
    )
    # optional bucket groups first token so variants like "tm us" and "tm one" stay near
    df_year["Type_bucket"] = (
        df_year["Type_key"].str.extract(r"^([a-z0-9]+)", expand=False)
        .fillna(df_year["Type_key"])
    )
    df_year = df_year.sort_values(
        ["Type_bucket", "Type_key", "Work_key", "Milestone Date"],
        kind="stable"
    )
    df_year["Group"] = df_year["Type"] + "\n(" + df_year["Workstream"] + ")"
    groups = list(dict.fromkeys(df_year["Group"]))  # preserve first-seen order

    # --- New slide ---
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # --- Legend across full width (top) ---
    legend_left   = LEFT_PAD
    legend_width  = prs.slide_width - LEFT_PAD - RIGHT_PAD
    slot_w        = legend_width / len(legend_items)
    legend_y      = legend_top

    for i, (label, shp_type, color) in enumerate(legend_items):
        x = legend_left + i * slot_w + (slot_w - Inches(0.3)) / 2
        m = slide.shapes.add_shape(shp_type, x, legend_y, Inches(0.3), Inches(0.3))
        m.fill.solid(); m.fill.fore_color.rgb = color
        m.line.fill.background()

        tb = slide.shapes.add_textbox(x + Inches(0.35), legend_y, slot_w - Inches(0.35), Inches(0.3))
        p  = tb.text_frame.paragraphs[0]
        p.text = label
        p.font.size = Pt(15)
        p.alignment = PP_ALIGN.LEFT

    # --- Chart geometry (grid area) ---
    chart_w  = prs.slide_width  - LEFT_PAD - RIGHT_PAD - sidebar_w
    chart_h  = prs.slide_height - top_margin - bottom_margin
    col_count = 12
    row_count = len(groups)
    col_w   = chart_w / col_count
    row_h   = chart_h / max(row_count, 1)

    left_origin = LEFT_PAD
    top_origin  = top_margin

    # --- Top month header band (across grid area only) ---
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left_origin + sidebar_w, top_origin,
        chart_w, header_h
    )
    header.fill.solid(); header.fill.fore_color.rgb = BLUE_HDR
    header.line.fill.solid(); header.line.fill.fore_color.rgb = CELL_BORDER
    header.line.width = Pt(0.95)

    # Month header cells and labels
    for m_idx, m in enumerate(pd.date_range(f"{year}-01-01", f"{year}-12-01", freq="MS")):
        cell = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left_origin + sidebar_w + m_idx * col_w, top_origin,
            col_w, header_h
        )
        cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_HDR
        cell.line.fill.solid(); cell.line.fill.fore_color.rgb = CELL_BORDER
        cell.line.width = Pt(0.95)

        tf = cell.text_frame
        tf.text = m.strftime("%b %y")
        p = tf.paragraphs[0]
        p.font.bold = True
        p.font.color.rgb = TEXT_WHITE
        p.font.size = Pt(18)
        p.alignment = PP_ALIGN.CENTER
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    # --- Sidebar header: Type / Workstream ---
    th1 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left_origin, top_origin, type_col_w, header_h
    )
    th1.fill.solid(); th1.fill.fore_color.rgb = BLUE_HDR
    th1.line.fill.solid(); th1.line.fill.fore_color.rgb = CELL_BORDER; th1.line.width = Pt(0.95)
    t1 = slide.shapes.add_textbox(left_origin, top_origin, type_col_w, header_h)
    p1 = t1.text_frame.paragraphs[0]
    p1.text = "Type"; p1.font.bold = True; p1.font.size = Pt(18); p1.font.color.rgb = TEXT_WHITE
    p1.alignment = PP_ALIGN.CENTER; t1.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    th2 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left_origin + type_col_w, top_origin, work_col_w, header_h
    )
    th2.fill.solid(); th2.fill.fore_color.rgb = BLUE_HDR
    th2.line.fill.solid(); th2.line.fill.fore_color.rgb = CELL_BORDER; th2.line.width = Pt(0.95)
    t2 = slide.shapes.add_textbox(left_origin + type_col_w, top_origin, work_col_w, header_h)
    p2 = t2.text_frame.paragraphs[0]
    p2.text = "Workstream"; p2.font.bold = True; p2.font.size = Pt(18); p2.font.color.rgb = TEXT_WHITE
    p2.alignment = PP_ALIGN.CENTER; t2.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # --- Sidebar rows + alternating banding & borders on date cells ---
    for r, grp in enumerate(groups):
        type_txt, work_txt = grp.split("\n")
        # Sidebar cells (Type & Workstream)
        for idx, (txt, x0, w0) in enumerate((
            (type_txt,                 left_origin,             type_col_w),
            (work_txt.strip("()"),     left_origin + type_col_w, work_col_w),
        )):
            cell = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                x0, top_origin + header_h + r * row_h,
                w0, row_h
            )
            # alternating shading: use the **date** grid colours or keep sidebar neutral
            cell.fill.solid(); cell.fill.fore_color.rgb = ROW_EVEN if r % 2 == 0 else ROW_ODD
            cell.line.fill.solid(); cell.line.fill.fore_color.rgb = CELL_BORDER; cell.line.width = Pt(0.95)

            tb = slide.shapes.add_textbox(x0, top_origin + header_h + r * row_h, w0, row_h)
            tf = tb.text_frame; tf.text = txt
            p  = tf.paragraphs[0]; p.font.size = Pt(15); p.font.color.rgb = TEXT_BLACK; p.alignment = PP_ALIGN.CENTER
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Date-grid cells with alternating banding + white borders
        for c in range(col_count):
            cell = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                left_origin + sidebar_w + c * col_w,
                top_origin + header_h + r * row_h,
                col_w, row_h
            )
            cell.fill.solid(); cell.fill.fore_color.rgb = ROW_EVEN if r % 2 == 0 else ROW_ODD
            cell.line.fill.solid(); cell.line.fill.fore_color.rgb = CELL_BORDER; cell.line.width = Pt(0.95)

        # Navy centerline through the row (single clean connector)
        y_center = top_origin + header_h + r * row_h + (row_h / 2)
        ln = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            left_origin + sidebar_w, y_center,
            left_origin + sidebar_w + chart_w, y_center
        )
        ln.line.fill.solid(); ln.line.fill.fore_color.rgb = ROW_CENTER; ln.line.width = Pt(0.5)

    # --- Milestones (accurate day-in-month positioning) ---
    for _, row in df_year.iterrows():
        dt        = row["Milestone Date"]
        month_idx = dt.month - 1
        dim       = monthrange(dt.year, dt.month)[1]
        day_frac  = (dt.day - 1) / float(dim - 1) if dim > 1 else 0.0  # 0..1

        # shape size & y-index
        is_major  = str(row.get("Milestone Type", "Regular")).strip().lower() == "major"
        size      = STAR_SIZE if is_major else CIRCLE_SIZE
        half      = size / 2.0
        y_idx     = groups.index(row["Group"])

        # positions
        x = left_origin + sidebar_w + (month_idx + day_frac) * col_w - half
        y = top_origin  + header_h   + y_idx * row_h + (row_h / 2.0) - half

        shp = slide.shapes.add_shape(SHAPE_MAP[row["Milestone Type"]], x, y, size, size)
        shp.fill.solid()
        shp.fill.fore_color.rgb = (
            legend_items[0][2] if is_major else status_colors.get(row["Milestone Status"], RGBColor(128,128,128))
        )
        shp.line.fill.background()

        # label to the right, slightly above
        lbl = slide.shapes.add_textbox(x + half + Inches(0.20), y - Inches(0.05), Inches(2.5), Inches(0.35))
        tf = lbl.text_frame
        tf.text = str(row["Milestone Title"])
        tf.paragraphs[0].font.size = Pt(12)
        tf.paragraphs[0].font.color.rgb = TEXT_BLACK

    # --- Today vertical green dotted line (only for this year) ---
    today = datetime.today().date()
    if today.year == year:
        dim   = monthrange(today.year, today.month)[1]
        tfrac = (today.day - 1) / float(dim - 1) if dim > 1 else 0.0
        xpos  = left_origin + sidebar_w + (today.month - 1 + tfrac) * col_w

        conn = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            xpos, top_origin + header_h,
            xpos, top_origin + header_h + chart_h
        )
        conn.line.fill.solid()
        conn.line.fill.fore_color.rgb = TODAY_GREEN
        conn.line.width = Pt(2)
        conn.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

# ========= 4) SAVE =========
OUT_PPTX = r"C:\Users\you\Desktop\Python Roadmap\Roadmap.pptx"
prs.save(OUT_PPTX)
print("Saved:", OUT_PPTX)