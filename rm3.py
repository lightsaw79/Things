# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE

import pandas as pd
from datetime import datetime, date
from calendar import monthrange

# ---------- 1) CONFIG ----------
EXCEL_PATH = r"C:\path\to\your\Input.xlsx"
OUT_PPTX   = r"C:\path\to\output\Roadmap.pptx"

SLIDE_W, SLIDE_H = Inches(20), Inches(9)   # keep your perfect layout
MAX_ROWS = 30                               # 30 rows per slide

# Geometry constants (kept same “look” you liked)
SIDEBAR_W   = Inches(4)                     # left table width
TYPE_COL_W  = Inches(1.5)                   # Type width (inside the table)
WORK_COL_W  = SIDEBAR_W - TYPE_COL_W
HEADER_H    = Inches(1.0)                   # top months header
LEGEND_H    = Inches(0.6)
LEGEND_TOP  = Inches(0.2)
TOP_MARGIN  = LEGEND_TOP + LEGEND_H + Inches(0.2)
BOTTOM_MARGIN = Inches(1.0)                 # space near bottom

LIGHT_COL   = RGBColor(224,242,255)
DARK_COL    = RGBColor(190,220,240)
BLUE_HDR    = RGBColor(91,155,213)
WHITE       = RGBColor(255,255,255)
NAVY        = RGBColor(0,0,128)
GREEN_DOT   = RGBColor(0,176,80)

# Status colors
STATUS_COLORS = {
    "On Track": RGBColor(0,176,80),
    "At Risk":  RGBColor(255,192,0),
    "Off Track":RGBColor(255,0,0),
    "Complete": RGBColor(0,112,192),
    "TBC":      RGBColor(191,191,191),
}

# Marker sizes
CIRCLE_SIZE = Inches(0.30)
STAR_SIZE   = Inches(0.40)

# ---------- 2) TABLE-CELL BORDER HELPER (fixes 'border_left' error) ----------
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

def set_cell_border(cell, *, left=None, right=None, top=None, bottom=None):
    """
    Apply borders to a python-pptx table cell via oxml.
    Each side spec is dict: {'sz': 635, 'color': 'FFFFFF', 'dash': 'solid'}
    sz in EMUs (12700 ≈ 1pt). 635 ≈ 0.05pt (hairline).
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    def _one_side(tag, spec):
        # remove existing side lines
        for el in tcPr.findall(qn(f"a:{tag}")):
            tcPr.remove(el)
        if not spec:
            return
        ln = OxmlElement(f"a:{tag}")                # a:lnL / a:lnR / a:lnT / a:lnB
        ln.set("w", str(spec.get("sz", 635)))

        solid = OxmlElement("a:solidFill")
        rgb = OxmlElement("a:srgbClr")
        rgb.set("val", spec.get("color", "FFFFFF"))
        solid.append(rgb)
        ln.append(solid)

        dash = OxmlElement("a:prstDash")
        dash.set("val", spec.get("dash", "solid"))
        ln.append(dash)

        ln.append(OxmlElement("a:round"))
        tcPr.append(ln)

    _one_side("lnL", left)
    _one_side("lnR", right)
    _one_side("lnT", top)
    _one_side("lnB", bottom)

BORDER_SPEC = {"sz": 12700, "color": "FFFFFF", "dash": "solid"}  # 1pt white grid borders


# ---------- 3) DATA ----------
df = pd.read_excel(EXCEL_PATH)
df["Milestone Date"] = pd.to_datetime(df["Milestone Date"])
df["Year"] = df["Milestone Date"].dt.year

# Cleaning for sort (stable/grouped alphabetical: Type, then Workstream)
def _norm(s):
    return (
        s.astype(str)
         .str.replace(r"\s+", " ", regex=True)
         .str.strip()
         .str.casefold()
    )

df["_type_key"] = _norm(df["Type"])
df["_work_key"] = _norm(df["Workstream"])

# Sort: Type (bucket+key), then Workstream, then date (stable keeps input order otherwise)
df = df.sort_values(by=["_type_key", "_work_key", "Milestone Date"], kind="stable")

# Group label used in the sidebar table and to align rows
df["Group"] = df["Type"].astype(str) + "\n(" + df["Workstream"].astype(str) + ")"


# ---------- 4) BUILD ONE SLIDE ----------
def build_slide(prs, subset_df, year, page_no, total_pages):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H

    # geometry
    chart_w = prs.slide_width - SIDEBAR_W - Inches(0.5)  # keep right margin ~0.5"
    chart_h = prs.slide_height - TOP_MARGIN - BOTTOM_MARGIN
    col_count = 12
    col_w = chart_w / col_count

    groups = list(dict.fromkeys(subset_df["Group"]))
    row_count = len(groups)
    row_h = chart_h / max(row_count, 1)

    LEFT_PAD  = Inches(0.25)
    left_origin = LEFT_PAD
    top_origin  = TOP_MARGIN

    # ---------- Legend (spread across top) ----------
    legend_items = [
        ("Major Milestone", MSO_SHAPE.STAR_5_POINT, RGBColor(0,176,80)),
        ("On Track",        MSO_SHAPE.OVAL,        STATUS_COLORS["On Track"]),
        ("At Risk",         MSO_SHAPE.OVAL,        STATUS_COLORS["At Risk"]),
        ("Off Track",       MSO_SHAPE.OVAL,        STATUS_COLORS["Off Track"]),
        ("Complete",        MSO_SHAPE.OVAL,        STATUS_COLORS["Complete"]),
        ("TBC",             MSO_SHAPE.OVAL,        STATUS_COLORS["TBC"]),
    ]
    legend_left = left_origin - LEFT_PAD
    legend_width = prs.slide_width - legend_left - Inches(0.5)
    slot_w = legend_width / len(legend_items)
    legend_y = Inches(0.2)

    for i, (label, shp_type, color) in enumerate(legend_items):
        x = legend_left + i*slot_w + (slot_w - Inches(0.3))/2
        mark = slide.shapes.add_shape(shp_type, x, legend_y, Inches(0.3), Inches(0.3))
        mark.fill.solid(); mark.fill.fore_color.rgb = color
        mark.line.fill.background()

        tb = slide.shapes.add_textbox(x + Inches(0.35), legend_y, Inches(1.0), Inches(0.3))
        p = tb.text_frame.paragraphs[0]
        p.text = label
        p.font.size = Pt(15)
        p.alignment = PP_ALIGN.LEFT

    # ---------- Header month row (blue blocks) ----------
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left_origin + SIDEBAR_W, top_origin,
        chart_w, HEADER_H
    )
    header.fill.solid(); header.fill.fore_color.rgb = BLUE_HDR
    header.line.fill.solid(); header.line.fill.fore_color.rgb = WHITE; header.line.width = Pt(0.95)

    for m, d in enumerate(pd.date_range(f"{year}-01-01", f"{year}-12-01", freq="MS")):
        # month cell
        cell = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left_origin + SIDEBAR_W + m*col_w, top_origin,
            col_w, HEADER_H
        )
        cell.fill.background()
        cell.line.fill.solid(); cell.line.fill.fore_color.rgb = WHITE; cell.line.width = Pt(0.95)

        # month text
        tb = slide.shapes.add_textbox(
            left_origin + SIDEBAR_W + m*col_w, top_origin,
            col_w, HEADER_H
        )
        p = tb.text_frame.paragraphs[0]
        p.text = d.strftime("%b %y")
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.font.size = Pt(18)
        p.alignment = PP_ALIGN.CENTER
        tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # ---------- Left sidebar as a REAL TABLE (resizable columns) ----------
    tbl_rows = int(row_count + 1)      # + header row
    tbl_cols = int(2)                  # Type, Workstream

    tbl_shape = slide.shapes.add_table(
        tbl_rows, tbl_cols,
        left_origin, top_origin,
        SIDEBAR_W, HEADER_H + row_count*row_h
    )
    tbl = tbl_shape.table

    # header cells
    tbl.columns[0].width = TYPE_COL_W
    tbl.columns[1].width = WORK_COL_W
    tbl.rows[0].height = HEADER_H
    for r in range(1, tbl_rows):
        tbl.rows[r].height = row_h

    # header text + border (white thin)
    for c, title in enumerate(("Type", "Workstream")):
        cell = tbl.cell(0, c)
        cell.text = title
        tf = cell.text_frame; p = tf.paragraphs[0]
        p.font.bold = True; p.font.size = Pt(18); p.alignment = PP_ALIGN.CENTER
        p.font.color.rgb = WHITE
        set_cell_border(cell, left=BORDER_SPEC, right=BORDER_SPEC, top=BORDER_SPEC, bottom=BORDER_SPEC)
        # fill header blue
        hdr_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left_origin + (TYPE_COL_W if c==1 else 0),
            top_origin,
            TYPE_COL_W if c==0 else WORK_COL_W,
            HEADER_H
        )
        hdr_shape.fill.solid(); hdr_shape.fill.fore_color.rgb = BLUE_HDR
        hdr_shape.line.fill.solid(); hdr_shape.line.fill.fore_color.rgb = WHITE; hdr_shape.line.width = Pt(0.95)
        # put header text in front
        slide.shapes._spTree.remove(hdr_shape._element)
        slide.shapes._spTree.insert(0, hdr_shape._element)

    # sidebar body rows (alternating shade + borders + text)
    for r, g in enumerate(groups, start=1):
        type_txt, work_txt = g.split("\n")
        for c, txt in enumerate((type_txt, work_txt.strip("()"))):
            cell = tbl.cell(r, c)
            cell.text = txt
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(15)
            p.alignment = PP_ALIGN.CENTER if c==0 else PP_ALIGN.LEFT
            p.font.color.rgb = RGBColor(0,0,0)
            # border
            set_cell_border(cell, left=BORDER_SPEC, right=BORDER_SPEC, top=BORDER_SPEC, bottom=BORDER_SPEC)
            # row shading under table (alternate)
            fill_rect = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                left_origin + (0 if c==0 else TYPE_COL_W),
                top_origin + HEADER_H + (r-1)*row_h,
                (TYPE_COL_W if c==0 else WORK_COL_W),
                row_h
            )
            fill_rect.fill.solid()
            fill_rect.fill.fore_color.rgb = LIGHT_COL if (r-1)%2==0 else DARK_COL
            fill_rect.line.fill.solid(); fill_rect.line.fill.fore_color.rgb = WHITE; fill_rect.line.width = Pt(0.95)
            # send background rect behind table
            slide.shapes._spTree.remove(fill_rect._element)
            slide.shapes._spTree.insert(0, fill_rect._element)

    # ---------- Body grid (date area) ----------
    # alternating columns across date grid
    for i in range(col_count):
        for r in range(row_count):
            rect = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                left_origin + SIDEBAR_W + i*col_w,
                top_origin + HEADER_H + r*row_h,
                col_w, row_h
            )
            rect.fill.solid()
            rect.fill.fore_color.rgb = LIGHT_COL if r%2==0 else DARK_COL
            rect.line.fill.solid()
            rect.line.fill.fore_color.rgb = WHITE
            rect.line.width = Pt(0.95)
            # send to back
            slide.shapes._spTree.remove(rect._element)
            slide.shapes._spTree.insert(0, rect._element)

    # navy mid-row guide
    for r in range(row_count):
        y_center = top_origin + HEADER_H + r*row_h + (row_h/2)
        ln = slide.shapes.add_shape(
            MSO_CONNECTOR.STRAIGHT,
            left_origin + SIDEBAR_W, y_center,
            left_origin + SIDEBAR_W + chart_w, y_center
        )
        ln.line.fill.solid()
        ln.line.fill.fore_color.rgb = NAVY
        ln.line.width = Pt(0.5)

    # ---------- Milestones (accurate within-month position) ----------
    shape_map = {"Regular": MSO_SHAPE.OVAL, "Major": MSO_SHAPE.STAR_5_POINT}

    # map group -> row index for Y position
    group_index = {g:i for i,g in enumerate(groups)}

    for _, row in subset_df.iterrows():
        dt = row["Milestone Date"]
        month_idx = dt.month - 1
        dim = monthrange(dt.year, dt.month)[1]
        day_frac = (dt.day - 1) / (dim - 1) if dim > 1 else 0.0

        # accurate X inside month
        x = left_origin + SIDEBAR_W + month_idx*col_w + day_frac*col_w
        # centered Y on its group row
        irow = group_index[row["Group"]]
        y = top_origin + HEADER_H + irow*row_h + row_h/2

        # choose shape & size
        shp_type = shape_map.get(str(row.get("Milestone Type","Regular")).strip().title(), MSO_SHAPE.OVAL)
        size = STAR_SIZE if shp_type == MSO_SHAPE.STAR_5_POINT else CIRCLE_SIZE
        # center the marker
        x_draw = x - size/2
        y_draw = y - size/2

        shp = slide.shapes.add_shape(shp_type, x_draw, y_draw, size, size)
        shp.fill.solid(); shp.fill.fore_color.rgb = STATUS_COLORS.get(str(row["Milestone Status"]), RGBColor(128,128,128))
        shp.line.color.rgb = RGBColor(0,0,0)

        # label to right & slightly above
        lbl = slide.shapes.add_textbox(x + Inches(0.1), y - Inches(0.12), Inches(2.2), Inches(0.35))
        p = lbl.text_frame.paragraphs[0]
        p.text = str(row["Milestone Title"])
        p.font.size = Pt(12)

    # ---------- Today dotted line (only for this year) ----------
    today = date.today()
    if today.year == year:
        days_in_month = monthrange(today.year, today.month)[1]
        month_idx = today.month - 1
        day_frac = (today.day - 1) / (days_in_month - 1) if days_in_month > 1 else 0.0
        xpos = left_origin + SIDEBAR_W + (month_idx + day_frac)*col_w
        conn = slide.shapes.add_shape(
            MSO_CONNECTOR.STRAIGHT,
            xpos, top_origin + HEADER_H,
            xpos, top_origin + HEADER_H + chart_h
        )
        conn.line.fill.solid()
        conn.line.fill.fore_color.rgb = GREEN_DOT
        conn.line.width = Pt(2)
        conn.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT

    # ---------- Footer small page tag ----------
    foot = slide.shapes.add_textbox(Inches(0.3), prs.slide_height - Inches(0.4), Inches(5), Inches(0.3))
    p = foot.text_frame.paragraphs[0]
    p.text = f"TM-US Roadmap — {year}  •  Page {page_no}/{total_pages}"
    p.font.size = Pt(10)
    p.alignment = PP_ALIGN.LEFT


# ---------- 5) PAGINATE BY YEAR, 30 PER SLIDE ----------
prs = Presentation()
prs.slide_width, prs.slide_height = SLIDE_W, SLIDE_H

for year in sorted(df["Year"].unique()):
    df_year = df[df["Year"] == year].copy()

    # Ordered groups for this year (already sorted above)
    groups_year = list(dict.fromkeys(df_year["Group"]))
    # chunk into pages
    pages = [groups_year[i:i+MAX_ROWS] for i in range(0, len(groups_year), MAX_ROWS)]
    total_pages = max(1, len(pages))

    for idx, group_slice in enumerate(pages, start=1):
        # -------- FIX for earlier '.isin' error: build a mask then index ----------
        mask = (df_year["Type"].astype(str) + "\n(" + df_year["Workstream"].astype(str) + ")").isin(group_slice)
        sub = df_year[mask].copy()

        # keep milestone rows in page group order (stable)
        order_index = {g:i for i,g in enumerate(group_slice)}
        sub["_gidx"] = sub["Group"].map(order_index)
        sub = sub.sort_values(by=["_gidx", "Milestone Date"], kind="stable").drop(columns="_gidx")

        build_slide(prs, sub, year, idx, total_pages)

# ---------- 6) SAVE ----------
prs.save(OUT_PPTX)
print("Saved:", OUT_PPTX)