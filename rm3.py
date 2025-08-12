# -*- coding: utf-8 -*-
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

# -----------------------
# CONFIG (layout & styles)
# -----------------------
SLIDE_W = Inches(20)
SLIDE_H = Inches(9)

TOP_PAD     = Inches(0.4)
LEGEND_H    = Inches(0.6)
TITLE_H     = Inches(0.4)
HEADER_H    = Inches(1.0)
BOTTOM_PAD  = Inches(0.09)
LEFT_PAD    = Inches(0.4)
RIGHT_PAD   = Inches(0.4)

SIDEBAR_W   = Inches(4.0)       # Type + Workstream block
TYPE_COL_W  = Inches(1.5)
WORK_COL_W  = SIDEBAR_W - TYPE_COL_W

COL_COUNT   = 12                # months per slide (Jan..Dec)

# colors
BLUE_HDR    = RGBColor(91,155,213)
MONTH_EVEN  = RGBColor(190,220,240)
MONTH_ODD   = RGBColor(224,242,255)
GRID_WHITE  = RGBColor(255,255,255)
NAVY_LINE   = RGBColor(0,0,128)
TODAY_GREEN = RGBColor(0,176,80)

STATUS_COLORS = {
    "On Track": RGBColor(0,176,80),
    "At Risk":  RGBColor(255,192,0),
    "Off Track":RGBColor(255,0,0),
    "Complete": RGBColor(0,112,192),
    "TBC":      RGBColor(191,191,191)
}

CIRCLE_SIZE = Inches(0.30)      # Regular milestone diameter
STAR_SIZE   = Inches(0.40)      # Major milestone

# --- NEW: label layout tuning ---
LABEL_W         = Inches(2.6)   # default label width (right-side layout)
LABEL_H_LINE    = Inches(0.18)  # approx line height for label text
LABEL_X_GAP     = Inches(0.10)  # gap from marker when placed to the right
LABEL_Y_OFFSET  = Inches(0.12)  # how far above/below the center line we start
CLUSTER_X_FRACT = 0.25          # two items closer than 25% of month width = cluster
LANES_PER_SIDE  = 3             # how many stacked lanes above/below to try
LAST_4_MONTHS_START_IDX = 8     # 0-based: 8..11 are Sep..Dec (last 4 months)

# -----------------------
# Helpers: sort & clean
# -----------------------
def clean_text(s: str) -> str:
    if pd.isna(s): return ""
    return str(s).strip()

def type_bucket(s: str) -> str:
    """
    Extracts leading word/letters+digits for grouping Types alphabetically
    but keeps like Types together (e.g., 'TM - Detection', 'TM US - One' → 'tm').
    """
    s = clean_text(s).casefold()
    m = re.match(r"([a-z0-9]+)", re.sub(r"[^a-z0-9]", "", s))
    return m.group(1) if m else s

# --- NEW: small text helpers ---
def wrap_by_words(text: str, n_words: int = 3) -> str:
    """Break text into lines every n words."""
    words = clean_text(text).split()
    if not words:
        return ""
    lines = [" ".join(words[i:i+n_words]) for i in range(0, len(words), n_words)]
    return "\n".join(lines)

def est_label_height(text: str) -> float:
    """Estimate label height in inches based on line count."""
    lines = clean_text(text).count("\n") + 1
    return float(lines) * float(LABEL_H_LINE)

# -----------------------
# Row height & pagination
# -----------------------
def get_row_height(row_count: int, neat_max=18, hard_max=26) -> float:
    if row_count <= neat_max:
        return Inches(0.40)
    if row_count <= hard_max:
        neat_h = 0.40
        min_h  = 0.28
        factor = (hard_max - row_count) / (hard_max - neat_max)
        return Inches(min_h + (neat_h - min_h) * factor)
    raise ValueError("Too many rows for one slide, paginate needed.")

def chunk_groups(groups, hard_max=26):
    out, cur = [], []
    for g in groups:
        cur.append(g)
        if len(cur) >= hard_max:
            out.append(cur)
            cur = []
    if cur: out.append(cur)
    return out

# -----------------------
# Geometry helpers
# -----------------------
def month_col_width():
    dates_area_w = SLIDE_W - LEFT_PAD - RIGHT_PAD - SIDEBAR_W
    return dates_area_w / COL_COUNT

def chart_height_for_rows(row_count, row_h):
    return row_count * row_h

def slide_top_origin():
    return TOP_PAD + LEGEND_H + TITLE_H

# -----------------------
# Drawing primitives
# -----------------------
def add_title(slide, text):
    tb = slide.shapes.add_textbox(LEFT_PAD, TOP_PAD, SLIDE_W - LEFT_PAD - RIGHT_PAD, TITLE_H)
    p = tb.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(24)
    p.font.bold = True

def add_legend(slide):
    items = [
        ("Major Milestone", MSO_SHAPE.STAR_5_POINT, STATUS_COLORS["On Track"]),
        ("On Track", MSO_SHAPE.OVAL, STATUS_COLORS["On Track"]),
        ("At Risk",  MSO_SHAPE.OVAL, STATUS_COLORS["At Risk"]),
        ("Off Track",MSO_SHAPE.OVAL, STATUS_COLORS["Off Track"]),
        ("Complete", MSO_SHAPE.OVAL, STATUS_COLORS["Complete"]),
        ("TBC",      MSO_SHAPE.OVAL, STATUS_COLORS["TBC"]),
    ]
    slot_w = (SLIDE_W - LEFT_PAD - RIGHT_PAD) / len(items)
    y = TOP_PAD + TITLE_H/4.0

    for i,(label, shp_type, color) in enumerate(items):
        x = LEFT_PAD + i*slot_w + Inches(0.2)
        s = slide.shapes.add_shape(shp_type, x, y, Inches(0.30), Inches(0.30))
        s.fill.solid(); s.fill.fore_color.rgb = color
        s.line.fill.background()

        tb = slide.shapes.add_textbox(x + Inches(0.4), y - Inches(0.02), Inches(2.5), Inches(0.4))
        p = tb.text_frame.paragraphs[0]
        p.text = label
        p.font.size = Pt(14)

def add_month_header(slide, year, left, top, col_w):
    hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, SIDEBAR_W + COL_COUNT*col_w, HEADER_H)
    hdr.fill.solid(); hdr.fill.fore_color.rgb = BLUE_HDR
    hdr.line.fill.background()

    tbox1 = slide.shapes.add_textbox(left, top, TYPE_COL_W, HEADER_H)
    p1 = tbox1.text_frame.paragraphs[0]; p1.text = "Type"
    p1.font.bold = True; p1.font.size = Pt(18); p1.alignment = PP_ALIGN.CENTER
    tbox1.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    tbox2 = slide.shapes.add_textbox(left + TYPE_COL_W, top, WORK_COL_W, HEADER_H)
    p2 = tbox2.text_frame.paragraphs[0]; p2.text = "Workstream"
    p2.font.bold = True; p2.font.size = Pt(18); p2.alignment = PP_ALIGN.CENTER
    tbox2.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    for i in range(COL_COUNT):
        mx = left + SIDEBAR_W + i*col_w
        cell = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, mx, top, col_w, HEADER_H)
        cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_HDR
        cell.line.fill.solid(); cell.line.fill.fore_color.rgb = GRID_WHITE; cell.line.width = Pt(0.95)

        lbl = slide.shapes.add_textbox(mx, top, col_w, HEADER_H)
        p = lbl.text_frame.paragraphs[0]
        d = date(year, i+1, 1)
        p.text = d.strftime("%b %y")
        p.font.color.rgb = GRID_WHITE
        p.font.bold = True
        p.font.size = Pt(18)
        p.alignment = PP_ALIGN.CENTER
        lbl.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

def add_sidebar_rows(slide, groups, left, top, row_h):
    for r, grp in enumerate(groups):
        y = top + r*row_h
        for (x, w, txt) in (
            (left, TYPE_COL_W, grp.split("\n")[0]),
            (left + TYPE_COL_W, WORK_COL_W, grp.split("\n")[1].strip("()"))
        ):
            cell = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, row_h)
            cell.fill.solid()
            cell.fill.fore_color.rgb = MONTH_ODD if (r % 2 == 0) else MONTH_EVEN
            cell.line.fill.solid(); cell.line.fill.fore_color.rgb = GRID_WHITE; cell.line.width = Pt(0.95)

            tb = slide.shapes.add_textbox(x, y, w, row_h)
            tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tb.text_frame.paragraphs[0]
            p.text = txt
            p.font.size = Pt(15)
            p.alignment = PP_ALIGN.CENTER

def add_dates_grid(slide, groups, left, top, col_w, row_h):
    row_count = len(groups)
    for i in range(COL_COUNT):
        for r in range(row_count):
            x = left + SIDEBAR_W + i*col_w
            y = top + r*row_h
            cell = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, col_w, row_h)
            cell.fill.solid()
            cell.fill.fore_color.rgb = MONTH_EVEN if (r % 2 == 0) else MONTH_ODD
            cell.line.fill.solid(); cell.line.fill.fore_color.rgb = GRID_WHITE; cell.line.width = Pt(0.95)

    x_left  = left + SIDEBAR_W
    x_right = left + SIDEBAR_W + COL_COUNT*col_w
    for r in range(row_count):
        y_center = top + r*row_h + row_h/2.0
        ln = slide.shapes.add_shape(MSO_CONNECTOR.STRAIGHT, x_left, y_center, x_right, y_center)
        ln.line.fill.solid(); ln.line.fill.fore_color.rgb = NAVY_LINE; ln.line.width = Pt(0.5)

def add_today_line_if_current_year(slide, year, left, top, col_w, row_h, row_count):
    today = date.today()
    if today.year != year:
        return
    days_in_month = monthrange(today.year, today.month)[1]
    day_frac = (today.day - 1) / (days_in_month - 1) if days_in_month > 1 else 0.0
    month_idx = today.month - 1
    xpos = left + SIDEBAR_W + (month_idx + day_frac) * col_w

    conn = slide.shapes.add_shape(
        MSO_CONNECTOR.STRAIGHT,
        xpos, top, xpos, top + row_count*row_h
    )
    conn.line.fill.solid()
    conn.line.fill.fore_color.rgb = TODAY_GREEN
    conn.line.width = Pt(2)
    conn.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

# --- NEW: geometry helpers for label collision ---
def rects_intersect(a, b):
    """a,b = (x0,y0,x1,y1) in inches; detect overlap."""
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    return (ax0 < bx1) and (bx0 < ax1) and (ay0 < by1) and (by0 < ay1)

def clamp(val, lo, hi):
    return max(lo, min(hi, val))

# -----------------------
# Collision‑aware milestones & labels
# -----------------------
def plot_milestones(slide, df_page, groups, year, left, top, col_w, row_h):
    """
    - Accurate x (day‑fraction inside month)
    - Collision‑aware labels per row (above/below lanes)
    - For last 4 months (Sep–Dec) or if the right‑side label would overflow,
      switch to centered‑above and wrap every 3 words
    """
    # per row bookkeeping
    row_boxes = [[] for _ in range(len(groups))]   # placed label rectangles to avoid overlaps
    row_last_x = [None]*len(groups)               # last marker x per row
    row_alt_above = [True]*len(groups)            # toggle above/below for clusters

    x_min = left + SIDEBAR_W
    x_max = x_min + COL_COUNT*col_w

    for _, row in df_page.iterrows():
        dt = pd.to_datetime(row["Milestone Date"]).date()
        if dt.year != year:
            continue

        month_idx = dt.month - 1
        dim = monthrange(dt.year, dt.month)[1]
        day_frac = (dt.day - 1) / (dim - 1) if dim > 1 else 0.0

        x = x_min + (month_idx + day_frac) * col_w
        grp_label = f"{row['Type']}\n({row['Workstream']})"
        try:
            y_index = groups.index(grp_label)
        except ValueError:
            continue
        y_center = top + y_index*row_h + row_h/2.0

        is_major = (str(row.get("Milestone Type","")).strip().casefold() == "major")
        size = STAR_SIZE if is_major else CIRCLE_SIZE
        half = size/2.0

        # milestone glyph
        shp = slide.shapes.add_shape(
            MSO_SHAPE.STAR_5_POINT if is_major else MSO_SHAPE.OVAL,
            x - half, y_center - half, size, size
        )
        status = clean_text(row.get("Milestone Status",""))
        shp.fill.solid(); shp.fill.fore_color.rgb = STATUS_COLORS.get(status, RGBColor(128,128,128))
        shp.line.color.rgb = RGBColor(0,0,0)

        # -------- smart label placement ----------
        raw_text = clean_text(row.get("Milestone Title",""))

        # Decide if right-side label would overflow; also for last 4 months we prefer above+wrapped.
        prefer_above = (month_idx >= LAST_4_MONTHS_START_IDX)
        if not prefer_above:
            # would the right-side box go out of the dates area?
            if x + LABEL_X_GAP + LABEL_W > x_max:
                prefer_above = True

        # If last 4 months: wrap every 3 words
        label_text = wrap_by_words(raw_text, 3) if prefer_above else raw_text
        label_h_est = est_label_height(label_text)

        # cluster detection with previous in the same row (push alternation)
        if row_last_x[y_index] is not None:
            if abs(x - row_last_x[y_index]) < (CLUSTER_X_FRACT * col_w):
                row_alt_above[y_index] = not row_alt_above[y_index]
        row_last_x[y_index] = x

        # initial candidate position
        if prefer_above:
            # centered above the marker, keep inside bounds
            bx = clamp(x - LABEL_W/2.0, x_min, x_max - LABEL_W)
            by = y_center - LABEL_Y_OFFSET - label_h_est
            # try lanes stacked upward then downward if needed
            lane_order = [0] + [i for k in range(1, LANES_PER_SIDE+1) for i in (k, -k)]
        else:
            # to the right of marker, slightly above/below alternately
            base_above = row_alt_above[y_index]
            by = y_center - (LABEL_Y_OFFSET if base_above else -LABEL_Y_OFFSET)
            bx = clamp(x + LABEL_X_GAP, x_min, x_max - LABEL_W)
            lane_order = [0] + [i for k in range(1, LANES_PER_SIDE+1) for i in ((-k) if base_above else k,
                                                                                (k) if base_above else -k)]

        # collision‑avoidance: shift by lanes until rectangle doesn’t intersect
        placed = None
        for lane in lane_order:
            # vertical shift per lane
            y_shift = lane * LABEL_H_LINE * 1.1
            rx0, ry0 = bx, by + y_shift
            rx1, ry1 = rx0 + LABEL_W, ry0 + label_h_est
            cand = (rx0, ry0, rx1, ry1)
            if all(not rects_intersect(cand, existing) for existing in row_boxes[y_index]):
                placed = cand
                row_boxes[y_index].append(cand)
                break

        if placed is None:
            # as a fallback, pin just above the glyph
            rx0 = clamp(x - LABEL_W/2.0, x_min, x_max - LABEL_W)
            ry0 = y_center - LABEL_Y_OFFSET - label_h_est
            placed = (rx0, ry0, rx0 + LABEL_W, ry0 + label_h_est)
            row_boxes[y_index].append(placed)

        # draw the textbox at 'placed'
        rx0, ry0, rx1, ry1 = placed
        lbl = slide.shapes.add_textbox(rx0, ry0, rx1 - rx0, ry1 - ry0)
        tf = lbl.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = label_text
        p.font.size = Pt(12)
        p.alignment = PP_ALIGN.LEFT if not prefer_above else PP_ALIGN.CENTER
        tf.vertical_anchor = MSO_ANCHOR.TOP

# -----------------------
# Slide builder (per year, per page)
# -----------------------
def add_title_and_legend(slide, title_prefix, year, page_num, total_pages):
    title_text = f"{title_prefix} — {year}"
    if total_pages > 1:
        title_text += f"  (Page {page_num}/{total_pages})"
    add_title(slide, title_text)
    add_legend(slide)

def build_slide(prs, df_page, year, page_num, total_pages, title_prefix="TM‑US Roadmap"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_title_and_legend(slide, title_prefix, year, page_num, total_pages)

    left = LEFT_PAD
    top  = slide_top_origin()
    col_w = month_col_width()

    groups = (
        df_page[["Type","Workstream"]]
        .drop_duplicates()
        .apply(lambda r: f"{r['Type']}\n({r['Workstream']})", axis=1)
        .tolist()
    )
    row_h = get_row_height(len(groups))
    chart_h = chart_height_for_rows(len(groups), row_h)

    add_month_header(slide, year, left, top, col_w)
    add_sidebar_rows(slide, groups, left, top + HEADER_H, row_h)
    add_dates_grid(slide, groups, left, top + HEADER_H, col_w, row_h)
    add_today_line_if_current_year(slide, year, left, top + HEADER_H, col_w, row_h, len(groups))
    plot_milestones(slide, df_page, groups, year, left, top + HEADER_H, col_w, row_h)

# -----------------------
# MAIN
# -----------------------
def build_ppt(input_xlsx: str, output_pptx: str, title_prefix="TM‑US Roadmap"):
    df = pd.read_excel(input_xlsx)

    # Normalize & sort
    df["Type"] = df["Type"].map(clean_text)
    df["Workstream"] = df["Workstream"].map(clean_text)
    df["Milestone Title"] = df["Milestone Title"].map(clean_text)
    df["Milestone Date"] = pd.to_datetime(df["Milestone Date"])
    df["year"] = df["Milestone Date"].dt.year

    df["Type_key"] = df["Type"].map(clean_text).str.casefold()
    df["Type_bucket"] = df["Type"].map(type_bucket)
    df["Work_key"] = df["Workstream"].map(clean_text).str.casefold()

    df = df.sort_values(by=["Type_bucket","Type_key","Work_key","Milestone Date"], kind="stable")

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    for year in sorted(df["year"].dropna().unique()):
        df_year = df[df["year"] == year].copy()

        groups = (
            df_year[["Type","Workstream","Type_bucket","Type_key","Work_key"]]
            .drop_duplicates()
            .sort_values(by=["Type_bucket","Type_key","Work_key"], kind="stable")
            .apply(lambda r: f"{r['Type']}\n({r['Workstream']})", axis=1)
            .tolist()
        )

        pages = chunk_groups(groups, hard_max=26)
        total_pages = len(pages)

        for page_num, page_groups in enumerate(pages, start=1):
            page_mask = df_year.apply(
                lambda r: f"{r['Type']}\n({r['Workstream']})" in page_groups, axis=1
            )
            df_page = df_year[page_mask].copy()

            build_slide(prs, df_page, year, page_num, total_pages, title_prefix)

    prs.save(output_pptx)

# -----------------------
# Run (example)
# -----------------------
if __name__ == "__main__":
    INPUT_XLSX = r"C:\path\to\your\Roadmap_Input_Sheet.xlsx"   # change me
    OUTPUT_PPTX = r"C:\path\to\your\Roadmap.pptx"              # change me
    build_ppt(INPUT_XLSX, OUTPUT_PPTX, title_prefix="TM‑US Roadmap")
    print("Saved:", OUTPUT_PPTX)