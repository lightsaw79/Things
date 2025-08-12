# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.dml.color import RGBColor

import pandas as pd
from datetime import date
from calendar import monthrange
import re

# =========================
# CONFIG / CONSTANTS
# =========================
SLIDE_W = Inches(20)
SLIDE_H = Inches(9)

LEFT_PAD_IN   = Inches(0.09)
RIGHT_PAD_IN  = Inches(0.09)

# ↓ tighter legend + padding so the table gets more space
TOP_PAD_IN    = Inches(0.18)   # was 0.40
LEGEND_H_IN   = Inches(0.45)   # was 0.60

HEADER_H      = Inches(1.00)
BOTTOM_PAD_IN = Inches(0.09)
MAX_ROWS_PER_SLIDE = 20

CIRCLE_SIZE = Inches(0.30)
STAR_SIZE   = Inches(0.40)

BLUE_HDR   = RGBColor(91,155,213)
MONTH_ODD  = RGBColor(224,242,255)
MONTH_EVEN = RGBColor(190,220,240)
WHITE      = RGBColor(255,255,255)
BLACK      = RGBColor(0,0,0)
NAVY       = RGBColor(0,0,128)
TODAY_GREEN= RGBColor(0,176,80)

status_colors = {
    "On Track":  RGBColor(0,176,80),
    "At Risk":   RGBColor(255,192,0),
    "Off Track": RGBColor(255,0,0),
    "Complete":  RGBColor(0,112,192),
    "TBC":       RGBColor(191,191,191),
}

# =========================
# HELPERS
# =========================
def clean(s: str) -> str:
    return str(s).strip().replace("\n"," ").casefold()

def type_bucket(s: str) -> str:
    cs = clean(s)
    m = re.match(r"([a-z0-9]+)", re.sub(r"[^a-z0-9]", "", cs))
    return m.group(1) if m else cs

def wrap_words(text: str, words_per_line: int = 3) -> str:
    words = re.findall(r"\S+", str(text))
    if not words: return ""
    return "\n".join(" ".join(words[i:i+words_per_line]) for i in range(0, len(words), words_per_line))

def bbox_intersects(a, b):
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    return not (ax2 <= bx1 or bx2 <= ax1 or ay2 <= by1 or by2 <= ay1)

def estimate_box(x_left, y_top, w_in, h_in):
    return (x_left, y_top, x_left + w_in, y_top + h_in)

def place_label_y(y_center_in, label_h_in, row_top_in, row_bottom_in, prefer="above"):
    pad = Inches(0.06)
    if prefer == "above":
        y_top = max(row_top_in + pad, y_center_in - label_h_in - pad)
        if y_top + label_h_in > row_bottom_in - pad:
            y_top = min(row_bottom_in - pad - label_h_in, y_center_in + pad)
    else:
        y_top = min(row_bottom_in - pad - label_h_in, y_center_in + pad)
        if y_top < row_top_in + pad:
            y_top = max(row_top_in + pad, y_center_in - label_h_in - pad)
    return y_top

def find_non_overlapping_y(proposed_box, placed_boxes, row_top_in, row_bottom_in):
    step = Inches(0.08)
    max_tries = 10
    for i in range(max_tries):
        if any(bbox_intersects(proposed_box, b) for b in placed_boxes):
            x1, y1, x2, y2 = proposed_box
            if i % 2 == 0:
                y1 = max(row_top_in + Inches(0.04), y1 - step)
            else:
                y1 = min(row_bottom_in - (y2 - y1) - Inches(0.04), y1 + step)
            proposed_box = (x1, y1, x2, y1 + (y2 - y1))
        else:
            break
    return proposed_box

def month_labels(year: int):
    return pd.date_range(f"{year}-01-01", f"{year}-12-01", freq="MS")

def add_full_width_legend(slide, left_in, top_in, width_in, height_in):
    items = [
        ("t0 Milestone", MSO_SHAPE.STAR_5_POINT, TODAY_GREEN),
        ("On Track",     MSO_SHAPE.OVAL,         status_colors["On Track"]),
        ("At Risk",      MSO_SHAPE.OVAL,         status_colors["At Risk"]),
        ("Off Track",    MSO_SHAPE.OVAL,         status_colors["Off Track"]),
        ("Complete",     MSO_SHAPE.OVAL,         status_colors["Complete"]),
        ("TBC",          MSO_SHAPE.OVAL,         status_colors["TBC"]),
    ]
    gap = Inches(0.18)  # a bit tighter
    slot = (width_in - gap*(len(items)-1)) / len(items)
    x = left_in
    for label, shp_type, color in items:
        # vertically centered inside the legend row
        m = slide.shapes.add_shape(shp_type, x, top_in + height_in/2 - Inches(0.13), Inches(0.26), Inches(0.26))
        m.fill.solid(); m.fill.fore_color.rgb = color
        m.line.color.rgb = BLACK
        tb = slide.shapes.add_textbox(x + Inches(0.34), top_in, slot - Inches(0.34), height_in)
        tf = tb.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE    # exact vertical center
        p = tf.paragraphs[0]
        p.text = label
        p.font.size = Pt(13)
        p.alignment = PP_ALIGN.LEFT
        x += slot + gap

def build_left_sidebar_table(slide, groups, left_in, top_in, work_col_w, type_col_w, row_h):
    rows = len(groups) + 1
    tbl_shape = slide.shapes.add_table(rows, 2, left_in, top_in, work_col_w + type_col_w, HEADER_H + row_h*len(groups))
    tbl = tbl_shape.table

    tbl.columns[0].width = type_col_w
    tbl.columns[1].width = work_col_w
    tbl.rows[0].height = HEADER_H
    for r in range(1, rows):
        tbl.rows[r].height = row_h

    # header cells — EXACT center (both axes)
    th_type = tbl.cell(0,0)
    th_work = tbl.cell(0,1)
    for th, label in ((th_type, "Type"), (th_work, "Workstream")):
        th.text_frame.clear()
        tf = th.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE        # vertical center
        p = tf.paragraphs[0]
        p.text = label
        p.font.bold = True
        p.font.size = Pt(18)
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER                 # horizontal center

    # white border rectangle to frame header
    hdr_border = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left_in, top_in, type_col_w + work_col_w, HEADER_H
    )
    hdr_border.fill.solid(); hdr_border.fill.fore_color.rgb = BLUE_HDR
    hdr_border.line.color.rgb = WHITE; hdr_border.line.width = Pt(1.0)
    hdr_border.z_order(1)   # send back

    # body cells — EXACT center
    for r, grp in enumerate(groups, start=1):
        t, w = grp.split("\n", 1)
        for cell, txt in ((tbl.cell(r,0), t), (tbl.cell(r,1), w)):
            tf = cell.text_frame
            tf.clear()
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.text = txt
            p.font.size = Pt(15)
            p.font.color.rgb = BLACK
            p.alignment = PP_ALIGN.CENTER

    return tbl, (top_in + HEADER_H)

def build_month_header_and_grid(slide, year, left_in, top_in, chart_w, row_h, row_count):
    col_w = chart_w / 12.0

    # month header
    for i, dt in enumerate(month_labels(year)):
        # bg rectangle
        cell = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_in + i*col_w, top_in, col_w, HEADER_H)
        cell.fill.solid(); cell.fill.fore_color.rgb = BLUE_HDR
        cell.line.color.rgb = WHITE; cell.line.width = Pt(0.95)

        # label textbox EXACT center (both axes)
        tb = slide.shapes.add_textbox(left_in + i*col_w, top_in, col_w, HEADER_H)
        tf = tb.text_frame; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = dt.strftime("%b %y")
        p.font.bold = True; p.font.size = Pt(18); p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    # grid cells + alternating row color
    grid_top = top_in + HEADER_H
    for i in range(12):
        for r in range(row_count):
            cell = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                left_in + i*col_w, grid_top + r*row_h,
                col_w, row_h
            )
            cell.fill.solid()
            cell.fill.fore_color.rgb = MONTH_ODD if r % 2 == 0 else MONTH_EVEN
            cell.line.color.rgb = WHITE
            cell.line.width = Pt(0.75)

    # thin navy line across each row (center)
    right_in = left_in + chart_w
    for r in range(row_count):
        y_center = grid_top + r*row_h + row_h/2.0
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, left_in, y_center, right_in, y_center)
        ln.line.color.rgb = NAVY
        ln.line.width = Pt(0.5)

    return col_w, right_in, grid_top

def add_today_line(slide, year, left_in, right_in, top_in, row_h, row_count, col_w):
    today = date.today()
    if today.year != year: return
    days_in = monthrange(today.year, today.month)[1]
    month_idx = today.month - 1
    day_frac = (today.day - 1) / (days_in - 1) if days_in > 1 else 0.0
    xpos = left_in + (month_idx + day_frac) * col_w
    grid_top = top_in + HEADER_H
    grid_bottom = grid_top + row_h * row_count
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, xpos, grid_top, xpos, grid_bottom)
    conn.line.color.rgb = TODAY_GREEN
    conn.line.width = Pt(2)
    conn.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

def build_groups_for_year(df_year):
    d = df_year.copy()
    d["Type_key"] = d["Type"].map(clean)
    d["Type_bucket"] = d["Type"].map(type_bucket)
    d["Work_key"] = d["Workstream"].map(clean)
    d = d.sort_values(by=["Type_bucket", "Type_key", "Work_key", "Milestone Date"], kind="stable")
    groups = (
        d[["Type", "Workstream"]]
        .drop_duplicates()
        .apply(lambda s: f"{s['Type']}\n{s['Workstream']}", axis=1)
        .tolist()
    )
    return groups, d

def plot_milestones(slide, df_rows, groups, left_in, right_in, grid_top, row_h, col_w, year):
    row_to_boxes = {r: [] for r in range(len(groups))}
    row_index = {g: i for i, g in enumerate(groups)}

    for _, row in df_rows.iterrows():
        # case-insensitive: t0/t1 OR major/regular (any capitalization)
        mtype_raw = str(row["Milestone Type"])
        mtype = mtype_raw.strip().lower()
        mtype = {"t0":"t0", "t1":"t1", "major":"t0", "regular":"t1"}.get(mtype, "t1")

        grp = f"{row['Type']}\n{row['Workstream']}"
        if grp not in row_index: continue
        r = row_index[grp]
        row_top = grid_top + r*row_h
        y_center = row_top + row_h/2.0

        dt = pd.to_datetime(row["Milestone Date"]).to_pydatetime()
        month_idx = dt.month - 1
        days_in_month = monthrange(dt.year, dt.month)[1]
        day_frac = (dt.day - 1)/(days_in_month - 1) if days_in_month > 1 else 0.0
        x_marker = left_in + (month_idx + day_frac) * col_w

        size = STAR_SIZE if mtype == "t0" else CIRCLE_SIZE
        shp = slide.shapes.add_shape(
            MSO_SHAPE.STAR_5_POINT if mtype == "t0" else MSO_SHAPE.OVAL,
            x_marker - size/2, y_center - size/2, size, size
        )
        shp.fill.solid()
        shp.fill.fore_color.rgb = status_colors.get(str(row["Milestone Status"]), RGBColor(191,191,191))
        shp.line.color.rgb = BLACK

        raw_text = str(row["Milestone Title"])
        is_last_four = month_idx >= 8
        label_text = wrap_words(raw_text, 3) if is_last_four else raw_text

        label_w = Inches(2.4 if not is_last_four else 2.2)
        label_h = Inches(0.55 if "\n" in label_text else 0.35)
        prefer_side = "above" if is_last_four else "below"

        if is_last_four:
            x_label = max(left_in, min(x_marker - label_w/2, right_in - label_w))
        else:
            x_label = min(x_marker + Inches(0.15), right_in - label_w)

        y_label = place_label_y(y_center, label_h, row_top, row_top + row_h, prefer=prefer_side)

        if not is_last_four:
            min_dx = Inches(0.10)
            if x_label - x_marker < min_dx:
                x_label = min(x_marker + min_dx, right_in - label_w)

        proposed = estimate_box(x_label, y_label, label_w, label_h)
        proposed = find_non_overlapping_y(proposed, row_to_boxes[r], row_top, row_top + row_h)

        lbl = slide.shapes.add_textbox(proposed[0], proposed[1], label_w, label_h)
        tf = lbl.text_frame
        tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE  # nicer vertical centering inside label box
        p = tf.paragraphs[0]
        p.text = label_text
        p.font.size = Pt(12)
        p.font.color.rgb = BLACK

        row_to_boxes[r].append(proposed)

def build_slide(prs, df_page, year, page_no, total_pages, show_legend=True):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    if show_legend:
        add_full_width_legend(slide, LEFT_PAD_IN, TOP_PAD_IN,
                              prs.slide_width - LEFT_PAD_IN - RIGHT_PAD_IN, LEGEND_H_IN)
    top_origin = TOP_PAD_IN + (LEGEND_H_IN if show_legend else 0)

    usable_h = SLIDE_H - top_origin - BOTTOM_PAD_IN
    row_count = df_page[["Type","Workstream"]].drop_duplicates().shape[0]
    row_h = (usable_h - HEADER_H) / max(1, row_count)

    left_origin = LEFT_PAD_IN
    right_limit = prs.slide_width - RIGHT_PAD_IN
    chart_w = right_limit - left_origin

    TYPE_COL_W = Inches(1.5)
    WORK_COL_W = Inches((SLIDE_W - LEFT_PAD_IN - RIGHT_PAD_IN) * 0.20) - TYPE_COL_W
    max_sidebar = (SLIDE_W - LEFT_PAD_IN - RIGHT_PAD_IN) * 0.30
    sidebar_w = min(TYPE_COL_W + WORK_COL_W, max_sidebar)

    groups, df_sorted = build_groups_for_year(df_page)

    tbl, first_body_top = build_left_sidebar_table(
        slide, groups, left_origin, top_origin, WORK_COL_W, TYPE_COL_W, row_h
    )

    months_left = left_origin + (TYPE_COL_W + WORK_COL_W)
    months_w = chart_w - (TYPE_COL_W + WORK_COL_W)
    col_w, right_in, grid_top = build_month_header_and_grid(
        slide, year, months_left, top_origin, months_w, row_h, len(groups)
    )

    add_today_line(slide, year, months_left, right_in, top_origin, row_h, len(groups), col_w)

    plot_milestones(slide, df_sorted, groups, months_left, right_in, grid_top, row_h, col_w, year)

def main():
    df = pd.read_excel(r"C:\path\to\your\Input.xlsx")

    df["Milestone Date"] = pd.to_datetime(df["Milestone Date"])
    df["year"] = df["Milestone Date"].dt.year
    df["Milestone Status"] = df["Milestone Status"].astype(str).str.strip().str.title()

    # case-insensitive milestone type mapping (handles 'Major', 'REGULAR', etc.)
    df["Milestone Type"] = (
        df["Milestone Type"].astype(str).str.strip().str.lower()
          .map({"t0":"t0", "t1":"t1", "major":"t0", "regular":"t1"})
          .fillna("t1")
    )

    df["Type_key"] = df["Type"].map(clean)
    df["Type_bucket"] = df["Type"].map(type_bucket)
    df["Work_key"] = df["Workstream"].map(clean)

    df = df.sort_values(by=["Type_bucket", "Type_key", "Work_key", "Milestone Date"], kind="stable")

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    for year in sorted(df["year"].dropna().unique().tolist()):
        df_year = df[df["year"] == year].copy()
        groups_all, df_year_sorted = build_groups_for_year(df_year)

        rows = groups_all
        idx = 0
        page_no = 1
        total_pages = (len(rows) + MAX_ROWS_PER_SLIDE - 1) // MAX_ROWS_PER_SLIDE or 1
        while idx < len(rows) or (len(rows) == 0 and page_no == 1):
            slice_rows = rows[idx: idx + MAX_ROWS_PER_SLIDE] if rows else []
            if slice_rows:
                sub = df_year_sorted.merge(
                    pd.DataFrame({"Type": [r.split("\n")[0] for r in slice_rows],
                                  "Workstream": [r.split("\n")[1] for r in slice_rows]}),
                    on=["Type","Workstream"], how="inner"
                )
            else:
                sub = df_year_sorted.iloc[0:0].copy()

            build_slide(prs, sub, year, page_no, total_pages, show_legend=True)
            idx += MAX_ROWS_PER_SLIDE
            page_no += 1

    OUT_PPTX = r"C:\path\to\output\Roadmap.pptx"
    prs.save(OUT_PPTX)
    print("Saved:", OUT_PPTX)

if __name__ == "__main__":
    main()