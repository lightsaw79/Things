# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

import pandas as pd
from datetime import date, datetime
from calendar import monthrange

# ---------------- CONFIG ----------------
INPUT_XLSX = r"C:\path\to\Roadmap_Input.xlsx"   # <-- change me
OUT_PPTX   = r"C:\path\to\Roadmap.pptx"         # <-- change me

SLIDE_W_IN, SLIDE_H_IN = 20, 9  # slide size in inches

LEFT_PAD   = Inches(0.5)
RIGHT_PAD  = Inches(0.5)
TOP_PAD    = Inches(1.4)   # space for legend
BOTTOM_PAD = Inches(0.5)

TYPE_COL_W = Inches(1.5)
WORK_COL_W = Inches(2.5)        # tweak if you want more/less sidebar width
HEADER_H   = Inches(1.0)
ROW_H_MIN  = Inches(0.6)        # minimum comfy row height (will shrink if needed)

COL_COUNT  = 12                 # months
ROWS_PER_PAGE = 30              # groups per slide

# Colors
BLUE_HDR   = RGBColor( 91,155,213)
ALT_ROW_1  = RGBColor(224,242,255)  # light
ALT_ROW_2  = RGBColor(190,220,240)  # medium
GRID_WHITE = RGBColor(255,255,255)
NAVY_LINE  = RGBColor(  0,  0,128)
TODAY_GRN  = RGBColor(  0,176, 80)
BLACK      = RGBColor(  0,  0,  0)
WHITE      = RGBColor(255,255,255)

# Status → colors
STATUS_COLORS = {
    "On Track": RGBColor(0,176,80),
    "At Risk":  RGBColor(255,192,0),
    "Off Track":RGBColor(255,0,0),
    "Complete": RGBColor(0,112,192),
    "TBC":      RGBColor(191,191,191),
}

# Milestone marker sizes and label offsets
CIRCLE_SIZE = Inches(0.30)
STAR_SIZE   = Inches(0.40)
LBL_DX      = Inches(0.40)      # label x offset to the right of marker
LBL_DY      = Inches(0.10)      # label y offset upward

# ---------------- HELPERS ----------------
def normalize_text(s):
    if pd.isna(s): return ""
    return " ".join(str(s).split()).casefold()

def type_bucket(s):
    s = normalize_text(s)
    if not s: return ""
    out = []
    for ch in s:
        if ch.isalnum() or ch in "-_/":
            out.append(ch)
        elif out:
            break
    return "".join(out)

def month_name(y, m):
    return date(y, m, 1).strftime("%b %y")

def day_fraction(d: date) -> float:
    dim = monthrange(d.year, d.month)[1]
    return 0.5 if dim <= 1 else (d.day - 1) / (dim - 1)

def add_textbox(slide, left, top, width, height, text, size=12,
                color=BLACK, bold=False, align=PP_ALIGN.LEFT,
                vanchor=MSO_ANCHOR.MIDDLE):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame; tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = align
    tf.vertical_anchor = vanchor
    return tb

def build_legend(slide):
    items = [
        ("Major Milestone", MSO_SHAPE.STAR_5_POINT, STATUS_COLORS["On Track"]),
        ("On Track",  MSO_SHAPE.OVAL, STATUS_COLORS["On Track"]),
        ("At Risk",   MSO_SHAPE.OVAL, STATUS_COLORS["At Risk"]),
        ("Off Track", MSO_SHAPE.OVAL, STATUS_COLORS["Off Track"]),
        ("Complete",  MSO_SHAPE.OVAL, STATUS_COLORS["Complete"]),
        ("TBC",       MSO_SHAPE.OVAL, STATUS_COLORS["TBC"]),
    ]
    band_left = LEFT_PAD
    band_w = Inches(SLIDE_W_IN) - LEFT_PAD - RIGHT_PAD
    slot = band_w / len(items)
    y = Inches(0.3)

    for i, (lbl, shp, col) in enumerate(items):
        x = band_left + slot*i + Inches(0.3)
        size = STAR_SIZE if "Major" in lbl else CIRCLE_SIZE
        s = slide.shapes.add_shape(shp, x, y, size, size)
        s.fill.solid(); s.fill.fore_color.rgb = col
        s.line.fill.solid(); s.line.fill.fore_color.rgb = BLACK; s.line.width = Pt(0.75)
        add_textbox(slide, x + size + Inches(0.15), y - Inches(0.02),
                    slot - (size + Inches(0.2)), size,
                    lbl, size=14, align=PP_ALIGN.LEFT)

# -------- FULL EDITABLE TABLE (Type, Workstream, Jan..Dec) --------
def build_full_editable_table(slide, year, groups, total_width, total_height):
    """
    Creates one table with:
      - Row 0: header [Type, Workstream, Jan..Dec]
      - Rows 1..N: body
    Returns:
      tbl_shape, tbl, month_cells (list of dict per month with left,width),
      row_cells (list with top,height per row), first_month_left, last_month_right
    """
    rows = int(len(groups) + 1)
    cols = int(2 + COL_COUNT)

    tbl_shape = slide.shapes.add_table(
        rows, cols,
        LEFT_PAD, TOP_PAD,
        total_width, total_height
    )
    tbl = tbl_shape.table
    tbl_shape.name = "RoadmapTable"  # handy name in case you add macros later

    # Column widths: set Type/Workstream fixed, distribute month area evenly.
    tbl.columns[0].width = TYPE_COL_W
    tbl.columns[1].width = WORK_COL_W
    month_area_w = total_width - TYPE_COL_W - WORK_COL_W
    base_w = int(month_area_w // COL_COUNT)
    remainder = int(month_area_w - base_w * COL_COUNT)
    for c in range(2, cols):
        w = base_w + (remainder if c == cols-1 else 0)  # put remainder into last month col
        tbl.columns[c].width = w

    # Row heights
    tbl.rows[0].height = HEADER_H
    body_h = total_height - HEADER_H
    row_count = max(1, len(groups))
    base_h = int(body_h // row_count)
    leftover = int(body_h - base_h * row_count)
    for r in range(1, rows):
        h = base_h + (leftover if r == rows-1 else 0)
        tbl.rows[r].height = h

    # Header row content + styling
    headers = ["Type", "Workstream"] + [month_name(year, m) for m in range(1, 13)]
    for c in range(cols):
        cell = tbl.cell(0, c)
        tf = cell.text_frame; tf.clear()
        p = tf.paragraphs[0]; p.text = headers[c]
        p.font.size = Pt(18); p.font.bold = True; p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_HDR

    # Body cells + alternating row shading for ALL columns
    for r, grp in enumerate(groups, start=1):
        shade = ALT_ROW_1 if (r % 2 == 1) else ALT_ROW_2
        type_txt, work_txt = grp.split("\n(", 1); work_txt = work_txt[:-1]
        for c in range(cols):
            cell = tbl.cell(r, c)
            tf = cell.text_frame; tf.clear()
            if c == 0:   # Type
                val = type_txt
            elif c == 1: # Workstream
                val = work_txt
            else:
                val = "" # month cells are empty background
            p = tf.paragraphs[0]; p.text = val
            p.font.size = Pt(15); p.font.color.rgb = BLACK
            p.alignment = PP_ALIGN.CENTER; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid(); cell.fill.fore_color.rgb = shade

    # Precompute geometry from table cells for precise shape placement
    # month_cells: index 0..11 => dict with 'left', 'width'
    month_cells = []
    for m in range(COL_COUNT):
        month_col = 2 + m
        left = tbl.cell(1, month_col).shape.left  # use first body row
        width = tbl.cell(1, month_col).shape.width
        month_cells.append({"left": left, "width": width})
    first_month_left  = month_cells[0]["left"]
    last_month_right  = month_cells[-1]["left"] + month_cells[-1]["width"]

    # row_cells: index 0..row_count-1 => dict with 'top', 'height'
    row_cells = []
    for r in range(len(groups)):
        top  = tbl.cell(r+1, 2).shape.top     # use month column (consistent top)
        height = tbl.cell(r+1, 2).shape.height
        row_cells.append({"top": top, "height": height})

    return tbl_shape, tbl, month_cells, row_cells, first_month_left, last_month_right

def draw_row_center_lines(slide, row_cells, first_month_left, last_month_right):
    """Thin navy line through the center of each row across month area."""
    for rc in row_cells:
        y = rc["top"] + rc["height"] / 2
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, first_month_left, y,
                                        last_month_right, y)
        ln.line.fill.solid(); ln.line.fill.fore_color.rgb = NAVY_LINE; ln.line.width = Pt(0.5)

def place_milestones_independent(slide, groups, subset_df, year, month_cells, row_cells):
    """Drop stars/circles + labels as independent shapes using the table geometry at generation time."""
    shape_map = {"Regular": MSO_SHAPE.OVAL, "Major": MSO_SHAPE.STAR_5_POINT}

    for _, r in subset_df.iterrows():
        dt = pd.to_datetime(r["Milestone Date"]).date()
        if dt.year != year:
            continue

        # locate the month cell and within-cell x (accurate by day)
        m_idx = dt.month - 1
        xf = day_fraction(dt)  # 0..1
        m_left = month_cells[m_idx]["left"]
        m_width = month_cells[m_idx]["width"]
        x_center = m_left + xf * m_width

        # locate the row center for this group
        group_label = f'{r["Type"]}\n({r["Workstream"]})'
        try:
            gi = groups.index(group_label)
        except ValueError:
            continue
        rc = row_cells[gi]
        y_center = rc["top"] + rc["height"] / 2

        # marker
        is_major = str(r.get("Milestone Type", "Regular")).strip().lower() == "major"
        size = STAR_SIZE if is_major else CIRCLE_SIZE
        half = size / 2.0
        shp = slide.shapes.add_shape(shape_map["Major" if is_major else "Regular"],
                                     x_center - half, y_center - half,
                                     size, size)
        shp.fill.solid()
        shp.fill.fore_color.rgb = STATUS_COLORS.get(r["Milestone Status"], RGBColor(128,128,128))
        shp.line.color.rgb = BLACK; shp.line.width = Pt(0.5)
        # IMPORTANT: no locking, no grouping, no tags — user can freely move/resize later

        # label (independent textbox)
        add_textbox(slide, x_center + size + LBL_DX, y_center - LBL_DY,
                    Inches(2.5), Inches(0.4), str(r["Milestone Title"]), size=12, align=PP_ALIGN.LEFT)

def add_today_line(slide, year, month_cells, row_cells):
    """Accurate current-date vertical dotted line (independent shape)."""
    today = datetime.today().date()
    if today.year != year:
        return
    dim = monthrange(today.year, today.month)[1]
    frac = 0.5 if dim <= 1 else (today.day - 1) / (dim - 1)
    m_idx = today.month - 1
    m_left = month_cells[m_idx]["left"]
    m_width = month_cells[m_idx]["width"]
    x = m_left + frac * m_width

    top = row_cells[0]["top"] if row_cells else TOP_PAD + HEADER_H
    bottom = row_cells[-1]["top"] + row_cells[-1]["height"] if row_cells else top
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x, top, x, bottom)
    ln.line.fill.solid(); ln.line.fill.fore_color.rgb = TODAY_GRN
    ln.line.width = Pt(2); ln.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

# -------------- build one slide --------------
def build_slide(prs, df_page, year, page_no, total_pages):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    build_legend(slide)

    # compute table size
    total_width  = Inches(SLIDE_W_IN) - LEFT_PAD - RIGHT_PAD
    total_height = Inches(SLIDE_H_IN) - TOP_PAD - BOTTOM_PAD

    groups = list(dict.fromkeys(df_page["Type"].astype(str) + "\n(" + df_page["Workstream"].astype(str) + ")"))
    row_count = max(1, len(groups))
    # keep rows within vertical space: distribute evenly, but don't exceed min row height too badly
    body_h = total_height - HEADER_H
    row_h = max(ROW_H_MIN * 0.6, min(ROW_H_MIN, body_h / row_count))  # soft clamp
    # recompute total height to match row_h integral layout
    total_height = HEADER_H + row_h * row_count

    # build one full editable table
    tbl_shape, tbl, month_cells, row_cells, first_month_left, last_month_right = build_full_editable_table(
        slide, year, groups, total_width, total_height
    )

    # navy guides
    draw_row_center_lines(slide, row_cells, first_month_left, last_month_right)

    # milestones (independent)
    place_milestones_independent(slide, groups, df_page, year, month_cells, row_cells)

    # today's line
    add_today_line(slide, year, month_cells, row_cells)

    # footer
    add_textbox(slide, LEFT_PAD, Inches(SLIDE_H_IN) - BOTTOM_PAD, Inches(10), Inches(0.3),
                f"Roadmap — {year}  (Page {page_no}/{total_pages})", size=10, color=RGBColor(80,80,80))

# -------------- main --------------
def main():
    df = pd.read_excel(INPUT_XLSX)

    # required columns
    needed = {"Type","Workstream","Milestone Title","Milestone Date","Milestone Status","Milestone Type"}
    missing = needed - set(df.columns)
    if missing:
        raise ValueError(f"Missing columns in Excel: {missing}")

    df["Milestone Date"] = pd.to_datetime(df["Milestone Date"], errors="coerce")
    df = df.dropna(subset=["Milestone Date"]).copy()
    df["Year"] = df["Milestone Date"].dt.year
    df["Group"] = df["Type"].astype(str) + "\n(" + df["Workstream"].astype(str) + ")"

    # sort: Type bucket → Type → Workstream → date
    df["type_key"] = df["Type"].astype(str).map(normalize_text)
    df["work_key"] = df["Workstream"].astype(str).map(normalize_text)
    df["type_bucket"] = df["Type"].astype(str).map(type_bucket)
    df = df.sort_values(by=["type_bucket", "type_key", "work_key", "Milestone Date"], kind="stable")

    prs = Presentation()
    prs.slide_width  = Inches(SLIDE_W_IN)
    prs.slide_height = Inches(SLIDE_H_IN)

    for year in sorted(df["Year"].dropna().unique()):
        df_year = df[df["Year"] == year].copy()

        # paginate by distinct groups (30 per slide)
        all_groups = list(dict.fromkeys(df_year["Group"]))
        pages = [all_groups[i:i+ROWS_PER_PAGE] for i in range(0, len(all_groups), ROWS_PER_PAGE)]
        total_pages = max(1, len(pages))

        if not pages:
            # build an empty scaffold slide so the year still appears
            build_slide(prs, df_year.head(0), year, 1, 1)
            continue

        for idx, group_slice in enumerate(pages, start=1):
            mask = df_year["Group"].isin(group_slice)
            sub = df_year[mask].copy()

            # keep milestones ordered by the page's group order, then by date
            order_index = {g:i for i,g in enumerate(group_slice)}
            sub["_gidx"] = sub["Group"].map(order_index)
            sub = sub.sort_values(by=["_gidx", "Milestone Date"], kind="stable").drop(columns="_gidx")

            build_slide(prs, sub, year, idx, total_pages)

    prs.save(OUT_PPTX)
    print("Saved:", OUT_PPTX)

if __name__ == "__main__":
    main()