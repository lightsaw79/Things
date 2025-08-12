# Roadmap → PowerPoint (python-pptx)
# Features: clean/sort; year→slides; 20 rows/slide with stretch-to-fit;
# left 2-column editable table; months grid; navy row center lines;
# accurate milestone x (day fraction); smart label staggering; Sep–Dec word-wrap;
# compact legend; white header text w/ border; today-line; margins = 0.09"

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import pandas as pd
import re
from datetime import datetime, date
from calendar import monthrange

# ---------------- Colors ----------------
BLUE_HDR = RGBColor(91,155,213)
WHITE    = RGBColor(255,255,255)
MONTH_ODD  = RGBColor(224,242,255)
MONTH_EVEN = RGBColor(190,220,240)
NAVY     = RGBColor(0,0,128)
GREEN    = RGBColor(0,176,80)
BLACK    = RGBColor(0,0,0)
STATUS_COLORS = {
    "on track": RGBColor(0,176,80),
    "at risk": RGBColor(255,192,0),
    "off track": RGBColor(255,0,0),
    "complete": RGBColor(0,112,192),
    "tbc": RGBColor(191,191,191),
}

# --------------- Layout (inches, floats) ---------------
SLIDE_W_IN  = 20.0
SLIDE_H_IN  = 9.0
LEFT_PAD_IN   = 0.09   # left margin
RIGHT_PAD_IN  = 0.09   # right margin
TOP_PAD_IN    = 0.12   # small top padding
LEGEND_H_IN   = 0.28   # compact legend height (reduced)
HEADER_H_IN   = 0.90   # month header band
BOTTOM_PAD_IN = 0.09

TYPE_COL_W_IN = 1.50
WORK_COL_W_IN = 4.00

MAX_ROWS_PER_SLIDE = 20          # hard limit per slide (you asked for 20)
MAJOR_SIZE_IN = 0.40             # star (T0)
REGULAR_SIZE_IN = 0.30           # circle (T1)
LABEL_FONT_PT = 12
MONTH_FONT_PT = 18
HDR_FONT_PT   = 18
BODY_FONT_PT  = 15

# --------------- Helpers ---------------
def clean_text(s: str) -> str:
    return str(s).strip().replace("\n", " ").casefold()

def type_bucket(s: str) -> str:
    s = clean_text(s)
    m = re.match(r"([a-z0-9]+)", re.sub(r"[^a-z0-9]", "", s))
    return m.group(1) if m else s

def wrap_every_n_words(text: str, n: int) -> str:
    words = text.split()
    out = []
    for i in range(0, len(words), n):
        out.append(" ".join(words[i:i+n]))
    return "\n".join(out)

def get_day_fraction(dt: date) -> float:
    days_in = monthrange(dt.year, dt.month)[1]
    # position with day precision, 0-based from month start
    return (dt.day - 1) / (days_in - 1) if days_in > 1 else 0.0

def compute_geometry_in(row_count: int):
    """All math in inches; convert to EMU only at API boundaries."""
    usable_w_in = SLIDE_W_IN - LEFT_PAD_IN - RIGHT_PAD_IN
    sidebar_w_in = TYPE_COL_W_IN + WORK_COL_W_IN
    dates_w_in   = usable_w_in - sidebar_w_in
    col_w_in     = dates_w_in / 12.0

    # Fill the vertical space down to BOTTOM_PAD_IN no matter how many rows (1..20)
    body_h_in = SLIDE_H_IN - TOP_PAD_IN - LEGEND_H_IN - BOTTOM_PAD_IN
    rows = max(1, min(MAX_ROWS_PER_SLIDE, row_count))
    row_h_in = (body_h_in - HEADER_H_IN) / rows

    left_in = LEFT_PAD_IN
    top_in  = TOP_PAD_IN + LEGEND_H_IN

    return dict(
        left_in=left_in, top_in=top_in,
        sidebar_w_in=sidebar_w_in,
        type_w_in=TYPE_COL_W_IN, work_w_in=WORK_COL_W_IN,
        dates_w_in=dates_w_in, col_w_in=col_w_in,
        header_h_in=HEADER_H_IN, row_h_in=row_h_in,
        total_w_in=usable_w_in, total_h_in=HEADER_H_IN + row_h_in * rows
    )

def build_legend(slide):
    """Compact legend, single row, top-left."""
    x = Inches(LEFT_PAD_IN)
    y = Inches(TOP_PAD_IN + 0.02)  # tiny inset
    # draw little bullet + text per status
    items = [
        ("On Track", STATUS_COLORS["on track"]),
        ("At Risk", STATUS_COLORS["at risk"]),
        ("Off Track", STATUS_COLORS["off track"]),
        ("Complete", STATUS_COLORS["complete"]),
        ("TBC", STATUS_COLORS["tbc"]),
    ]
    cursor_x = LEFT_PAD_IN
    for label, color in items:
        shp = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(cursor_x), y, Inches(0.18), Inches(0.18)
        )
        shp.fill.solid()
        shp.fill.fore_color.rgb = color
        shp.line.color.rgb = BLACK
        shp.line.width = Pt(0.5)

        tb = slide.shapes.add_textbox(
            Inches(cursor_x + 0.24), y, Inches(1.2), Inches(0.22)
        )
        p = tb.text_frame.paragraphs[0]
        p.text = label
        p.font.size = Pt(12)
        p.font.color.rgb = BLACK
        cursor_x += 1.6  # move right for next item

def build_left_table(slide, groups, geom):
    """Editable Type + Workstream table, with zebra body and white header text."""
    rows = int(len(groups) + 1)
    cols = 2
    total_h_in = geom["header_h_in"] + geom["row_h_in"] * len(groups)
    total_w_in = geom["type_w_in"] + geom["work_w_in"]
    left_in = geom["left_in"]
    top_in  = geom["top_in"]

    tbl_shape = slide.shapes.add_table(
        rows, cols,
        Inches(left_in), Inches(top_in),
        Inches(total_w_in), Inches(total_h_in)
    )
    tbl = tbl_shape.table

    # Set columns
    tbl.columns[0].width = Inches(geom["type_w_in"])
    tbl.columns[1].width = Inches(geom["work_w_in"])

    # Set row heights (header + body)
    tbl.rows[0].height = Inches(geom["header_h_in"])
    for r in range(1, rows):
        tbl.rows[r].height = Inches(geom["row_h_in"])

    # Header cells
    hdrs = ("Type", "Workstream")
    for c, title in enumerate(hdrs):
        cell = tbl.cell(0, c)
        cell.text = title
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(HDR_FONT_PT)
        p.font.bold = True
        p.font.color.rgb = WHITE
        # header background
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE_HDR

    # Body cells zebra + center text
    for r in range(1, rows):
        for c in range(2):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = MONTH_ODD if (r % 2 == 1) else MONTH_EVEN
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(BODY_FONT_PT)
            p.font.color.rgb = BLACK

    # outer border
    tbl_shape.line.color.rgb = WHITE
    tbl_shape.line.width = Pt(0.75)

    # Populate body (centered vertically)
    for r, (typ, work) in enumerate(groups, start=1):
        tbl.cell(r, 0).text = str(typ)
        tbl.cell(r, 1).text = str(work)
        for c in range(2):
            tf = tbl.cell(r, c).text_frame
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER

    return tbl_shape, tbl

def build_month_header_and_grid(slide, groups, geom, year):
    """Draw month header row (rectangles) and the 12xN grid rectangles only for dates area."""
    left_in  = geom["left_in"] + geom["sidebar_w_in"]
    top_in   = geom["top_in"]
    col_w_in = geom["col_w_in"]
    row_h_in = geom["row_h_in"]

    # Month header cells
    for m in range(12):
        cell = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(left_in + m * col_w_in), Inches(top_in),
            Inches(col_w_in), Inches(geom["header_h_in"])
        )
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE_HDR
        cell.line.color.rgb = WHITE
        cell.line.width = Pt(0.5)
        tb = slide.shapes.add_textbox(
            Inches(left_in + m * col_w_in), Inches(top_in),
            Inches(col_w_in), Inches(geom["header_h_in"])
        )
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        month_lbl = date(year, m+1, 1).strftime("%b %y")
        p.text = month_lbl
        p.font.size = Pt(MONTH_FONT_PT)
        p.font.color.rgb = WHITE
        p.font.bold = True

    # Month grid (dates area only)
    for m in range(12):
        for r in range(len(groups)):
            y = geom["top_in"] + geom["header_h_in"] + r * row_h_in
            rect = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(left_in + m * col_w_in), Inches(y),
                Inches(col_w_in), Inches(row_h_in)
            )
            rect.fill.solid()
            rect.fill.fore_color.rgb = MONTH_ODD if (r % 2 == 0) else MONTH_EVEN
            rect.line.color.rgb = WHITE
            rect.line.width = Pt(0.5)

    # Navy horizontal center lines across dates area (one per row)
    x_left  = left_in
    x_right = left_in + 12 * col_w_in
    for r in range(len(groups)):
        y_c = geom["top_in"] + geom["header_h_in"] + r * row_h_in + row_h_in/2.0
        ln = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(x_left), Inches(y_c),
            Inches(x_right), Inches(y_c)
        )
        ln.line.color.rgb = NAVY
        ln.line.width = Pt(0.5)

def build_groups_for_year(df_year_sorted):
    """Return list of (Type, Workstream) in required (stable) order."""
    # Already sorted/cleaned before calling; just take unique preserving order
    seen = set()
    out = []
    for _, row in df_year_sorted.iterrows():
        key = (row["Type"], row["Workstream"])
        if key not in seen:
            seen.add(key)
            out.append(key)
    return out

def place_milestones_and_labels(slide, df_page, groups, geom, year):
    """Draw milestones (T0 star / T1 circle) and smart labels."""
    left_in = geom["left_in"] + geom["sidebar_w_in"]
    col_w_in = geom["col_w_in"]
    row_h_in = geom["row_h_in"]

    # Prepare per-row last label end x to help spacing
    last_label_end_x = {g: None for g in groups}
    # For last 4 months (Sep-Dec)
    last_month_start = 9

    for _, row in df_page.iterrows():
        dt = row["Milestone Date"].date() if isinstance(row["Milestone Date"], pd.Timestamp) else row["Milestone Date"]
        month_idx = dt.month - 1
        day_frac  = get_day_fraction(dt)
        x_in = left_in + (month_idx + day_frac) * col_w_in

        # row index by group
        grp_key = (row["Type"], row["Workstream"])
        if grp_key not in groups:
            continue
        r_index = groups.index(grp_key)
        y_mid = geom["top_in"] + geom["header_h_in"] + r_index * row_h_in + row_h_in/2.0

        # choose shape by type (case-insensitive mapping for 't0'/'major' vs 't1'/'regular')
        tval = str(row.get("Milestone Type", "")).strip().casefold()
        is_major = (tval == "t0") or (tval == "major")
        size_in = MAJOR_SIZE_IN if is_major else REGULAR_SIZE_IN
        half = size_in/2.0

        shp = slide.shapes.add_shape(
            MSO_SHAPE.STAR_5_POINT if is_major else MSO_SHAPE.OVAL,
            Inches(x_in - half), Inches(y_mid - half),
            Inches(size_in), Inches(size_in)
        )
        # fill and line
        status = str(row.get("Milestone Status","")).casefold()
        shp.fill.solid()
        shp.fill.fore_color.rgb = STATUS_COLORS.get(status, RGBColor(0,176,80))
        shp.line.color.rgb = BLACK
        shp.line.width = Pt(0.5)

        # label positioning rules
        label_text = str(row.get("Milestone Title","")).strip()

        # For Sep–Dec, wrap every 3 words and place ABOVE to avoid running off slide
        if month_idx >= (last_month_start-1):
            label_text = wrap_every_n_words(label_text, 3)
            place_above = True
        else:
            # smart staggering: if too close to previous label in same row, alternate above/below
            last_end = last_label_end_x.get(grp_key)
            min_gap_in = 0.35  # ~0.35" minimal spacing
            place_above = False
            if last_end is not None and (x_in - last_end) < min_gap_in:
                place_above = True

        # label box
        tb_w_in = 2.0  # starting width; we won't let it overflow right edge
        tb_h_in = 0.6
        # clamp x so box stays inside dates area
        right_limit_in = left_in + 12 * col_w_in
        x_label_in = min(x_in + half + 0.10, right_limit_in - tb_w_in)

        y_offset = - (half + 0.08) if place_above else (half + 0.08)
        tb = slide.shapes.add_textbox(
            Inches(x_label_in),
            Inches(y_mid + y_offset),
            Inches(tb_w_in),
            Inches(tb_h_in)
        )
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = label_text
        p.font.size = Pt(LABEL_FONT_PT)
        p.font.color.rgb = BLACK
        p.alignment = PP_ALIGN.LEFT

        # update last end for this row
        approx_char_w_in = 0.09  # rough average @ 12pt (tunable)
        est_len_in = len(label_text.replace("\n","")) * approx_char_w_in
        last_label_end_x[grp_key] = x_in + est_len_in

def draw_today_line(slide, geom, year):
    """Vertical dotted green line for current date if it matches slide year."""
    today = datetime.today().date()
    if today.year != year:
        return
    left_in  = geom["left_in"] + geom["sidebar_w_in"]
    col_w_in = geom["col_w_in"]
    month_idx = today.month - 1
    day_frac  = get_day_fraction(today)
    x_in = left_in + (month_idx + day_frac) * col_w_in
    y_top_in = geom["top_in"] + geom["header_h_in"]
    y_bot_in = y_top_in + geom["row_h_in"] * MAX_ROWS_PER_SLIDE  # long enough; below is fine

    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(x_in), Inches(y_top_in),
        Inches(x_in), Inches(y_bot_in)
    )
    conn.line.color.rgb = GREEN
    conn.line.width = Pt(2)
    conn.line.dash_style = 2  # round dot

# --------------- Main ---------------
def main():
    # ---- Load and prep data ----
    # Expected columns: Type, Workstream, Milestone Title, Milestone Date, Milestone Status, Milestone Type
    # Replace path below
    EXCEL_PATH = r"C:\path\to\your\Roadmap_Input.xlsx"
    df = pd.read_excel(EXCEL_PATH)

    # Normalize / clean
    df["Milestone Date"] = pd.to_datetime(df["Milestone Date"])
    df["year"] = df["Milestone Date"].dt.year

    df["Type_key"] = df["Type"].map(clean_text)
    df["Type_bucket"] = df["Type"].map(type_bucket)
    df["Work_key"] = df["Workstream"].map(clean_text)

    # Stable sort: bucket → type → work → date
    df = df.sort_values(by=["Type_bucket", "Type_key", "Work_key", "Milestone Date"], kind="stable")

    prs = Presentation()
    prs.slide_width  = Inches(SLIDE_W_IN)
    prs.slide_height = Inches(SLIDE_H_IN)

    # iterate years (ascending)
    for year in sorted(df["year"].dropna().unique()):
        df_year = df[df["year"] == year].copy()

        # Build the display order of unique (Type, Workstream)
        groups_sorted = build_groups_for_year(df_year)

        # Paginate 20 per slide
        total_rows = len(groups_sorted)
        pages = (total_rows + MAX_ROWS_PER_SLIDE - 1) // MAX_ROWS_PER_SLIDE
        if pages == 0:
            # still create an empty slide with header/legend for the year
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            build_legend(slide)
            geom = compute_geometry_in(1)
            build_left_table(slide, [], geom)  # empty
            build_month_header_and_grid(slide, [], geom, year)
            draw_today_line(slide, geom, year)
            continue

        for page_no in range(pages):
            start = page_no * MAX_ROWS_PER_SLIDE
            end   = min(start + MAX_ROWS_PER_SLIDE, total_rows)
            groups_page = groups_sorted[start:end]

            # Geometry depends on row count on THIS page (stretches to fill)
            geom = compute_geometry_in(len(groups_page))
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # Legend (compact)
            build_legend(slide)

            # Left table (editable) with zebra fill and centered text
            tbl_shape, tbl = build_left_table(slide, groups_page, geom)

            # Month header & grid in dates area
            build_month_header_and_grid(slide, groups_page, geom, year)

            # Subset milestones for just these groups
            mask = df_year.apply(lambda r: (r["Type"], r["Workstream"]) in set(groups_page), axis=1)
            df_page = df_year[mask].copy()

            # Draw milestones + labels (smart staggering + Sep–Dec wrap)
            place_milestones_and_labels(slide, df_page, groups_page, geom, year)

            # Today vertical line (if same year)
            draw_today_line(slide, geom, year)

    OUT_PPTX = r"C:\path\to\output\Roadmap.pptx"
    prs.save(OUT_PPTX)
    print("Saved:", OUT_PPTX)

if __name__ == "__main__":
    main()