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
SHEET_NAME = 0

df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
df["Milestone Date"] = pd.to_datetime(df["Milestone Date"])
df["Year"] = df["Milestone Date"].dt.year
df["Group"] = df["Type"] + "\n(" + df["Workstream"] + ")"

# ========= 2) CONSTANTS (same look/feel you liked) =========
prs = Presentation()
prs.slide_width  = Inches(20)
prs.slide_height = Inches(9)

# Colors
BLUE_HDR   = RGBColor( 91,155,213)
TEXT_WHITE = RGBColor(255,255,255)
TEXT_BLACK = RGBColor(  0,  0,  0)
ROW_EVEN   = RGBColor(224,242,255)  # light blue
ROW_ODD    = RGBColor(190,220,240)  # medium blue
CELL_BORDER= RGBColor(255,255,255)
ROW_CENTER = RGBColor(  0,  0,128)  # navy line mid-row
TODAY_GREEN= RGBColor(  0,176, 80)

status_colors = {
    "On Track":   RGBColor(144,238,144),
    "At Risk":    RGBColor(255,192,  0),
    "Off Track":  RGBColor(255,  0,  0),
    "Complete":   RGBColor(  0,112,192),
    "TBC":        RGBColor(191,191,191),
}

legend_items = [
    ("Major Milestone", MSO_SHAPE.STAR_5_POINT, RGBColor(192,164, 72)),
    ("On Track",        MSO_SHAPE.OVAL,         status_colors["On Track"]),
    ("At Risk",         MSO_SHAPE.OVAL,         status_colors["At Risk"]),
    ("Off Track",       MSO_SHAPE.OVAL,         status_colors["Off Track"]),
    ("Complete",        MSO_SHAPE.OVAL,         status_colors["Complete"]),
    ("TBC",             MSO_SHAPE.OVAL,         status_colors["TBC"]),
]

# Geometry
sidebar_w     = Inches(4)        # table width = Type + Workstream
type_col_w    = Inches(1.5)
work_col_w    = sidebar_w - type_col_w
header_h      = Inches(1.0)
legend_h      = Inches(0.6)
legend_top    = Inches(0.2)
LEFT_PAD      = Inches(0.25)
RIGHT_PAD     = Inches(0.25)
top_margin    = legend_top + legend_h + Inches(0.2)
bottom_margin = Inches(0.5)

# Symbol sizes
CIRCLE_SIZE = Inches(0.30)
STAR_SIZE   = Inches(0.40)
SHAPE_MAP   = {"Regular": MSO_SHAPE.OVAL, "Major": MSO_SHAPE.STAR_5_POINT}

# ========= 3) PER-YEAR SLIDES =========
for year in sorted(df["Year"].unique()):
    df_year = df[df["Year"] == year].copy()

    # Sort: Type A→Z (bucket keeps close variants together), then Workstream
    df_year["Type_key"] = (
        df_year["Type"].fillna("")
        .str.replace(r"\s+", " ", regex=True).str.strip().str.casefold()
    )
    df_year["Work_key"] = (
        df_year["Workstream"].fillna("")
        .str.replace(r"\s+", " ", regex=True).str.strip().str.casefold()
    )
    df_year["Type_bucket"] = (
        df_year["Type_key"].str.extract(r"^([a-z0-9]+)", expand=False)
        .fillna(df_year["Type_key"])
    )
    df_year = df_year.sort_values(
        ["Type_bucket", "Type_key", "Work_key", "Milestone Date"],
        kind="stable"
    )
    df_year["Group"] = df_year["Type"] + "\n(" + df_year["Workstream"] + ")"
    groups = list(dict.fromkeys(df_year["Group"]))

    # New slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Legend (full width)
    legend_left  = LEFT_PAD
    legend_width = prs.slide_width - LEFT_PAD - RIGHT_PAD
    slot_w       = legend_width / len(legend_items)
    for i, (label, shp_type, color) in enumerate(legend_items):
        x = legend_left + i * slot_w + (slot_w - Inches(0.3)) / 2
        y = legend_top
        m = slide.shapes.add_shape(shp_type, x, y, Inches(0.3), Inches(0.3))
        m.fill.solid(); m.fill.fore_color.rgb = color
        m.line.fill.background()
        tb = slide.shapes.add_textbox(x + Inches(0.35), y, slot_w - Inches(0.35), Inches(0.3))
        p  = tb.text_frame.paragraphs[0]
        p.text = label; p.font.size = Pt(15); p.alignment = PP_ALIGN.LEFT

    # Grid geometry
    chart_w  = prs.slide_width  - LEFT_PAD - RIGHT_PAD - sidebar_w
    chart_h  = prs.slide_height - top_margin - bottom_margin
    col_count= 12
    row_count= max(len(groups), 1)
    col_w    = chart_w / col_count
    row_h    = chart_h / row_count
    left_origin = LEFT_PAD
    top_origin  = top_margin

    # --------- A) LEFT SIDEBAR AS A REAL TABLE (resizable columns) ---------
    tbl_rows = row_count + 1  # +1 for header row
    tbl_cols = 2              # Type, Workstream
    tbl_shape = slide.shapes.add_table(
        tbl_rows, tbl_cols,
        left_origin, top_origin,
        sidebar_w, header_h + row_count * row_h
    )
    tbl = tbl_shape.table
    # column widths (resizable by user later)
    tbl.columns[0].width = type_col_w
    tbl.columns[1].width = work_col_w
    # row heights
    tbl.rows[0].height = header_h
    for r in range(1, tbl_rows):
        tbl.rows[r].height = row_h

    # Header cells
    for c, title in enumerate(("Type", "Workstream")):
        cell = tbl.cell(0, c)
        cell.text = title
        tf = cell.text_frame; p = tf.paragraphs[0]
        p.font.bold = True; p.font.size = Pt(18); p.font.color.rgb = TEXT_WHITE
        p.alignment = PP_ALIGN.CENTER
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        # blue fill
        cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_HDR

    # Body cells + alternating fill to match the date grid
    for r, grp in enumerate(groups, start=1):
        type_txt, work_txt = grp.split("\n")
        for c, text in enumerate((type_txt, work_txt.strip("()"))):
            cell = tbl.cell(r, c)
            cell.text = text
            tf = cell.text_frame; p = tf.paragraphs[0]
            p.font.size = Pt(15); p.font.color.rgb = TEXT_BLACK
            p.alignment = PP_ALIGN.CENTER
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            # banding
            base = ROW_EVEN if (r-1) % 2 == 0 else ROW_ODD
            cell.fill.solid(); cell.fill.fore_color.rgb = base

    # --------- B) MONTH HEADER (rectangles, as before) ---------
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left_origin + sidebar_w, top_origin,
        chart_w, header_h
    )
    header.fill.solid(); header.fill.fore_color.rgb = BLUE_HDR
    header.line.fill.solid(); header.line.fill.fore_color.rgb = CELL_BORDER; header.line.width = Pt(0.95)

    for m_idx, m in enumerate(pd.date_range(f"{year}-01-01", f"{year}-12-01", freq="MS")):
        cell = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left_origin + sidebar_w + m_idx * col_w, top_origin,
            col_w, header_h
        )
        cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_HDR
        cell.line.fill.solid(); cell.line.fill.fore_color.rgb = CELL_BORDER; cell.line.width = Pt(0.95)
        tf = cell.text_frame; tf.text = m.strftime("%b %y")
        p = tf.paragraphs[0]; p.font.bold = True; p.font.size = Pt(18); p.font.color.rgb = TEXT_WHITE
        p.alignment = PP_ALIGN.CENTER; tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    # --------- C) DATE GRID (rectangles + borders + navy centerline) ---------
    for r, grp in enumerate(groups):
        base = ROW_EVEN if r % 2 == 0 else ROW_ODD
        # date cells
        for c in range(col_count):
            cell = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                left_origin + sidebar_w + c * col_w,
                top_origin + header_h + r * row_h,
                col_w, row_h
            )
            cell.fill.solid(); cell.fill.fore_color.rgb = base
            cell.line.fill.solid(); cell.line.fill.fore_color.rgb = CELL_BORDER; cell.line.width = Pt(0.95)

        # row centerline
        y_center = top_origin + header_h + r * row_h + (row_h / 2)
        ln = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            left_origin + sidebar_w, y_center,
            left_origin + sidebar_w + chart_w, y_center
        )
        ln.line.fill.solid(); ln.line.fill.fore_color.rgb = ROW_CENTER; ln.line.width = Pt(0.5)

    # --------- D) MILESTONES (accurate day-in-month placement) ---------
    for _, row in df_year.iterrows():
        dt        = row["Milestone Date"]
        month_idx = dt.month - 1
        dim       = monthrange(dt.year, dt.month)[1]
        day_frac  = (dt.day - 1) / float(dim - 1) if dim > 1 else 0.0

        is_major  = str(row.get("Milestone Type", "Regular")).strip().lower() == "major"
        size      = STAR_SIZE if is_major else CIRCLE_SIZE
        half      = size / 2.0
        y_idx     = groups.index(row["Group"])

        x = left_origin + sidebar_w + (month_idx + day_frac) * col_w - half
        y = top_origin  + header_h + y_idx * row_h + (row_h / 2.0) - half

        shp = slide.shapes.add_shape(SHAPE_MAP[row["Milestone Type"]], x, y, size, size)
        shp.fill.solid()
        shp.fill.fore_color.rgb = (RGBColor(192,164,72) if is_major
                                   else status_colors.get(row["Milestone Status"], RGBColor(128,128,128)))
        shp.line.fill.background()

        # label to the right, slightly above
        lbl = slide.shapes.add_textbox(x + half + Inches(0.20), y - Inches(0.05), Inches(2.5), Inches(0.35))
        tf = lbl.text_frame; tf.text = str(row["Milestone Title"])
        tf.paragraphs[0].font.size = Pt(12); tf.paragraphs[0].font.color.rgb = TEXT_BLACK

    # --------- E) TODAY LINE (only on matching year) ---------
    today = datetime.today().date()
    if today.year == year:
        dim   = monthrange(today.year, today.month)[1]
        tfrac = (today.day - 1) / float(dim - 1) if dim > 1 else 0.0
        xpos  = left_origin + sidebar_w + (today.month - 1 + tfrac) * col_w
        v = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            xpos, top_origin + header_h,
            xpos, top_origin + header_h + row_count * row_h
        )
        v.line.fill.solid(); v.line.fill.fore_color.rgb = TODAY_GREEN
        v.line.width = Pt(2); v.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

# ========= 4) SAVE =========
OUT_PPTX = r"C:\Users\you\Desktop\Python Roadmap\Roadmap.pptx"
prs.save(OUT_PPTX)
print("Saved:", OUT_PPTX)