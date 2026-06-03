from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import ChartData
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ── Colors ──────────────────────────────────────────────────────────────────
C_PRIMARY_BLUE  = RGBColor(0x00, 0x66, 0xCC)
C_FOCUS_BLUE    = RGBColor(0x00, 0x71, 0xE3)
C_SKY_BLUE      = RGBColor(0x29, 0x97, 0xFF)
C_WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
C_PARCHMENT     = RGBColor(0xF5, 0xF5, 0xF7)
C_NEAR_BLACK    = RGBColor(0x1D, 0x1D, 0x1F)
C_DARK_TILE1    = RGBColor(0x27, 0x27, 0x29)
C_DARK_TILE2    = RGBColor(0x2A, 0x2A, 0x2C)
C_PURE_BLACK    = RGBColor(0x00, 0x00, 0x00)
C_BODY_MUTED    = RGBColor(0xCC, 0xCC, 0xCC)
C_HAIRLINE      = RGBColor(0xE0, 0xE0, 0xE0)
C_INK_MUTED80   = RGBColor(0x33, 0x33, 0x33)

FONT = "Calibri"

# ── Presentation setup ───────────────────────────────────────────────────────
prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK_LAYOUT = prs.slide_layouts[6]   # completely blank

W  = prs.slide_width
H  = prs.slide_height
WI = 13.33   # inches
HI = 7.5

# ── Helpers ──────────────────────────────────────────────────────────────────

def set_bg(slide, rgb: RGBColor):
    """Fill slide background with a solid colour."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = rgb


def add_rect(slide, left, top, width, height, fill_rgb=None, line_rgb=None, line_width_pt=0.5):
    from pptx.util import Pt as _Pt
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    if fill_rgb:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_rgb
    else:
        shape.fill.background()
    if line_rgb:
        shape.line.color.rgb = line_rgb
        shape.line.width = Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape


def add_rounded_rect(slide, left, top, width, height, fill_rgb, adj=0.1, line_rgb=None):
    """Add a rounded rectangle (pill / card shape)."""
    from pptx.util import Pt as _Pt
    shape = slide.shapes.add_shape(
        5,  # ROUNDED_RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    # Set corner radius via XML adjustment
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    prstGeom = spPr.find(qn('a:prstGeom'))
    if prstGeom is not None:
        avLst = prstGeom.find(qn('a:avLst'))
        if avLst is None:
            avLst = etree.SubElement(prstGeom, qn('a:avLst'))
        else:
            for gd in avLst.findall(qn('a:gd')):
                avLst.remove(gd)
        gd = etree.SubElement(avLst, qn('a:gd'))
        gd.set('name', 'adj')
        # adj value in EMU-like units: 50000 = 50% = fully rounded (pill)
        gd.set('fmla', f'val {int(adj * 100000)}')

    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if line_rgb:
        shape.line.color.rgb = line_rgb
        shape.line.width = Pt(0.5)
    else:
        shape.line.fill.background()
    return shape


def add_text_box(slide, text, left, top, width, height,
                 font_size=12, bold=False, color=C_NEAR_BLACK,
                 align=PP_ALIGN.LEFT, font_name=FONT, wrap=True):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox


def add_text_box_multiline(slide, lines, left, top, width, height,
                           font_size=12, bold=False, color=C_NEAR_BLACK,
                           align=PP_ALIGN.LEFT, space_before_pt=4):
    """Add a textbox with multiple paragraphs."""
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(space_before_pt)
        run = p.add_run()
        run.text = line
        run.font.name = FONT
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = color
    return txBox


def pill_button(slide, text, cx, cy, w=1.8, h=0.38):
    """Centred pill CTA button."""
    shape = add_rounded_rect(slide, cx - w/2, cy - h/2, w, h, C_PRIMARY_BLUE, adj=0.5)
    tf = shape.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(10)
    run.font.bold = True
    run.font.color.rgb = C_WHITE
    return shape


def set_chart_font(chart, font_name=FONT, font_size_pt=10, color=None):
    """Apply font to chart title and axis labels."""
    try:
        if chart.has_title:
            tf = chart.chart_title.text_frame
            for para in tf.paragraphs:
                for run in para.runs:
                    run.font.name = font_name
                    run.font.size = Pt(font_size_pt)
    except Exception:
        pass


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 1 – Title Slide Light
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_WHITE)

add_text_box(slide,
    "Introducing the Future of Design",
    1.0, 1.8, 11.33, 1.2,
    font_size=40, bold=True, color=C_NEAR_BLACK, align=PP_ALIGN.CENTER)

add_text_box(slide,
    "A comprehensive design language built for clarity, depth, and human connection.",
    1.5, 3.1, 10.33, 0.8,
    font_size=20, bold=False, color=C_NEAR_BLACK, align=PP_ALIGN.CENTER)

# Two pill CTAs
pill_button(slide, "Learn More",      cx=5.17, cy=5.5)
pill_button(slide, "Get Started →",   cx=8.16, cy=5.5)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 2 – Title Slide Dark
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_DARK_TILE1)

add_text_box(slide,
    "Designed for Everyone",
    1.0, 2.0, 11.33, 1.2,
    font_size=40, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

add_text_box(slide,
    "Elegant, intentional, and built to last.",
    1.5, 3.3, 10.33, 0.8,
    font_size=20, bold=False, color=C_BODY_MUTED, align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 3 – Section Divider
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_DARK_TILE1)

add_text_box(slide, "01",
    0.7, 1.4, 3.0, 1.2,
    font_size=56, bold=True, color=C_PRIMARY_BLUE, align=PP_ALIGN.LEFT)

add_text_box(slide, "Design Foundations",
    0.7, 2.7, 8.0, 1.0,
    font_size=28, bold=True, color=C_WHITE, align=PP_ALIGN.LEFT)

add_text_box(slide,
    "Principles that guide every pixel and every interaction.",
    0.7, 3.7, 8.0, 0.7,
    font_size=15, bold=False, color=C_BODY_MUTED, align=PP_ALIGN.LEFT)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 4 – Content + Body (60/40 split)
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_PARCHMENT)

# Left text area (60%)
add_text_box(slide, "Clarity Above All Else",
    0.6, 0.9, 7.4, 0.8,
    font_size=24, bold=True, color=C_NEAR_BLACK)

body_text = (
    "Every element on screen earns its place. Negative space breathes life into "
    "layouts, while careful typographic hierarchy guides the eye naturally from "
    "headline to supporting detail.\n\n"
    "Consistent colour application reinforces meaning—blue for action, white for "
    "openness, dark tiles for focus—so users always know where they are and what "
    "they can do next."
)
add_text_box(slide, body_text,
    0.6, 1.85, 7.2, 4.5,
    font_size=12, bold=False, color=C_NEAR_BLACK, wrap=True)

# Right image placeholder (40%) – rounded rect
img_ph = add_rounded_rect(slide,
    left=8.3, top=0.7, width=4.4, height=5.9,
    fill_rgb=C_HAIRLINE, adj=0.05)
img_ph.line.color.rgb = C_HAIRLINE
img_ph.line.width = Pt(0.75)
# Label
add_text_box(slide, "Image / Media",
    8.3, 3.3, 4.4, 0.5,
    font_size=10, bold=False, color=RGBColor(0x99,0x99,0x99), align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 5 – Full Bullets
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_WHITE)

add_text_box(slide, "Key Principles",
    0.7, 0.5, 11.93, 0.8,
    font_size=24, bold=True, color=C_NEAR_BLACK)

# Hairline rule under heading
add_rect(slide, 0.7, 1.35, 11.93, 0.02, fill_rgb=C_HAIRLINE)

bullets = [
    "Hierarchy — every element has a clear visual rank that guides the reader effortlessly.",
    "Consistency — repeated patterns build familiarity and reduce cognitive load.",
    "Accessibility — minimum 4.5 : 1 contrast ratios ensure legibility for all users.",
    "Purposeful Motion — animations clarify state changes rather than decorate them.",
    "Economy — restraint in colour, type, and decoration lets content take centre stage.",
]

bullet_top = 1.55
for i, bullet in enumerate(bullets):
    # Blue square bullet marker
    add_rect(slide, 0.72, bullet_top + i*1.0 + 0.08, 0.12, 0.12, fill_rgb=C_PRIMARY_BLUE)
    add_text_box(slide, bullet,
        1.0, bullet_top + i*1.0, 11.6, 0.85,
        font_size=12, bold=False, color=C_NEAR_BLACK, wrap=True)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 6 – Bar Chart
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_PARCHMENT)

add_text_box(slide, "Quarterly Performance",
    0.6, 0.3, 12.13, 0.7,
    font_size=24, bold=True, color=C_NEAR_BLACK)

chart_data = ChartData()
chart_data.categories = ['Q1', 'Q2', 'Q3', 'Q4']
chart_data.add_series('Product Revenue',  (42, 58, 67, 81))
chart_data.add_series('Service Revenue',  (18, 24, 31, 40))

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(0.6), Inches(1.2), Inches(12.13), Inches(5.8),
    chart_data
).chart

chart.has_legend = True
chart.legend.position = 2   # bottom

# Colour series
from pptx.oxml.ns import nsmap
series0 = chart.series[0]
series1 = chart.series[1]

def set_series_color(series, rgb: RGBColor):
    spPr = series._element.find('.//' + qn('c:spPr'))
    if spPr is None:
        spPr = etree.SubElement(series._element, qn('c:spPr'))
    solidFill = spPr.find(qn('a:solidFill'))
    if solidFill is not None:
        spPr.remove(solidFill)
    solidFill = etree.SubElement(spPr, qn('a:solidFill'))
    srgbClr   = etree.SubElement(solidFill, qn('a:srgbClr'))
    srgbClr.set('val', f'{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}')

set_series_color(series0, C_PRIMARY_BLUE)
set_series_color(series1, C_SKY_BLUE)

# Plot area background – white card (via XML on the chart element)
try:
    plotArea = chart._element.find('.//' + qn('c:plotArea'))
    if plotArea is not None:
        spPr = plotArea.find(qn('c:spPr'))
        if spPr is None:
            spPr = etree.SubElement(plotArea, qn('c:spPr'))
        solidFill = etree.SubElement(spPr, qn('a:solidFill'))
        srgbClr   = etree.SubElement(solidFill, qn('a:srgbClr'))
        srgbClr.set('val', 'FFFFFF')
except Exception:
    pass


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 7 – Line Chart
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_WHITE)

add_text_box(slide, "Monthly Active Users",
    0.6, 0.3, 12.13, 0.7,
    font_size=24, bold=True, color=C_NEAR_BLACK)

chart_data = ChartData()
chart_data.categories = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
chart_data.add_series('Mobile',  (12, 19, 23, 31, 38, 45))
chart_data.add_series('Desktop', (20, 22, 25, 27, 30, 34))

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.LINE,
    Inches(0.6), Inches(1.2), Inches(12.13), Inches(5.8),
    chart_data
).chart

chart.has_legend = True
chart.legend.position = 2

def set_line_series_color(series, rgb: RGBColor):
    """Set line colour for a line-chart series."""
    spPr = series._element.find(qn('c:spPr'))
    if spPr is None:
        spPr = etree.SubElement(series._element, qn('c:spPr'))
    ln = spPr.find(qn('a:ln'))
    if ln is None:
        ln = etree.SubElement(spPr, qn('a:ln'))
    ln.set('w', '25400')   # 2 pt
    solidFill = ln.find(qn('a:solidFill'))
    if solidFill is not None:
        ln.remove(solidFill)
    solidFill = etree.SubElement(ln, qn('a:solidFill'))
    srgbClr   = etree.SubElement(solidFill, qn('a:srgbClr'))
    srgbClr.set('val', f'{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}')

set_line_series_color(chart.series[0], C_PRIMARY_BLUE)
set_line_series_color(chart.series[1], C_FOCUS_BLUE)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 8 – Pie Chart (dark bg)
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_DARK_TILE1)

add_text_box(slide, "Revenue by Segment",
    0.6, 0.3, 12.13, 0.7,
    font_size=24, bold=True, color=C_WHITE)

chart_data = ChartData()
chart_data.categories = ['Hardware', 'Software', 'Services', 'Licensing', 'Other']
chart_data.add_series('Share', (35, 25, 20, 12, 8))

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.PIE,
    Inches(1.5), Inches(1.2), Inches(10.33), Inches(5.8),
    chart_data
).chart

chart.has_legend = True
chart.legend.position = 2

# Colour individual pie slices
pie_colors = [C_PRIMARY_BLUE, C_SKY_BLUE, C_FOCUS_BLUE, C_BODY_MUTED, C_PARCHMENT]
plot = chart.plots[0]
for idx, point in enumerate(plot.series[0].points):
    pPt = point._element
    spPr = pPt.find(qn('c:spPr'))
    if spPr is None:
        spPr = etree.SubElement(pPt, qn('c:spPr'))
    solidFill = etree.SubElement(spPr, qn('a:solidFill'))
    rgb = pie_colors[idx % len(pie_colors)]
    srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
    srgbClr.set('val', f'{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}')


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 9 – Data Table
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_WHITE)

add_text_box(slide, "Product Comparison",
    0.6, 0.3, 12.13, 0.7,
    font_size=24, bold=True, color=C_NEAR_BLACK)

rows, cols = 5, 4
table_shape = slide.shapes.add_table(
    rows, cols,
    Inches(0.6), Inches(1.2), Inches(12.13), Inches(5.5)
)
tbl = table_shape.table

headers = ["Feature", "Starter", "Professional", "Enterprise"]
data = [
    ["Storage",       "10 GB",    "100 GB",   "Unlimited"],
    ["Users",         "1",        "10",        "Unlimited"],
    ["Support",       "Email",    "Priority",  "Dedicated"],
    ["API Access",    "Limited",  "Full",      "Full + SLA"],
]

for c_idx, hdr in enumerate(headers):
    cell = tbl.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_NEAR_BLACK
    tf = cell.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = hdr
    run.font.name = FONT
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = C_WHITE

for r_idx, row_data in enumerate(data):
    bg = C_WHITE if r_idx % 2 == 0 else C_PARCHMENT
    for c_idx, val in enumerate(row_data):
        cell = tbl.cell(r_idx + 1, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
        tf = cell.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = tf.paragraphs[0].add_run()
        run.text = val
        run.font.name = FONT
        run.font.size = Pt(12)
        run.font.color.rgb = C_NEAR_BLACK

# Set borders via XML
def set_cell_border(cell, rgb=C_HAIRLINE, width_pt=0.5):
    tc = cell._tc
    tcPr = tc.find(qn('a:tcPr'))
    if tcPr is None:
        tcPr = etree.SubElement(tc, qn('a:tcPr'))
    w_emu = int(width_pt * 12700)
    for edge in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
        ln = etree.SubElement(tcPr, qn(edge))
        ln.set('w', str(w_emu))
        ln.set('cap', 'flat')
        ln.set('cmpd', 'sng')
        solidFill = etree.SubElement(ln, qn('a:solidFill'))
        srgbClr   = etree.SubElement(solidFill, qn('a:srgbClr'))
        srgbClr.set('val', f'{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}')

for r in range(rows):
    for c in range(cols):
        set_cell_border(tbl.cell(r, c))


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 10 – Quote / Callout
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_DARK_TILE1)

# Decorative quote mark
add_text_box(slide, "❝",
    0.5, 0.3, 3.0, 1.5,
    font_size=72, bold=False, color=C_PRIMARY_BLUE, align=PP_ALIGN.LEFT)

add_text_box(slide,
    "Design is not just what it looks like and feels like.\nDesign is how it works.",
    1.0, 1.8, 11.33, 2.0,
    font_size=24, bold=False, color=C_WHITE, align=PP_ALIGN.CENTER)

add_text_box(slide, "— Steve Jobs",
    1.0, 3.9, 11.33, 0.6,
    font_size=14, bold=False, color=C_BODY_MUTED, align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 11 – Two-Column Comparison
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_WHITE)

add_text_box(slide, "Light vs. Dark Experience",
    0.5, 0.3, 12.33, 0.7,
    font_size=24, bold=True, color=C_NEAR_BLACK)

GAP = 0.15
COL_W = (WI - 1.0 - GAP) / 2   # ≈ 6.09

# Left column – parchment
left_box = add_rect(slide, 0.5, 1.2, COL_W, 5.8, fill_rgb=C_PARCHMENT)
left_box.line.fill.background()

add_text_box(slide, "Light Mode",
    0.7, 1.45, COL_W - 0.4, 0.5,
    font_size=15, bold=True, color=C_NEAR_BLACK)

left_items = [
    "Optimised for bright, natural environments",
    "High-contrast text on white backgrounds",
    "Ideal for document and reading tasks",
]
for i, item in enumerate(left_items):
    add_text_box(slide, f"• {item}",
        0.7, 2.15 + i*0.85, COL_W - 0.4, 0.75,
        font_size=12, color=C_NEAR_BLACK, wrap=True)

# Right column – dark
right_x = 0.5 + COL_W + GAP
right_box = add_rect(slide, right_x, 1.2, COL_W, 5.8, fill_rgb=C_DARK_TILE1)
right_box.line.fill.background()

add_text_box(slide, "Dark Mode",
    right_x + 0.2, 1.45, COL_W - 0.4, 0.5,
    font_size=15, bold=True, color=C_WHITE)

right_items = [
    "Reduces eye strain in low-light settings",
    "Vibrant accent colours pop on dark canvas",
    "Battery-friendly on OLED displays",
]
for i, item in enumerate(right_items):
    add_text_box(slide, f"• {item}",
        right_x + 0.2, 2.15 + i*0.85, COL_W - 0.4, 0.75,
        font_size=12, color=C_WHITE, wrap=True)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE 12 – Closing Slide
# ════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(BLANK_LAYOUT)
set_bg(slide, C_PURE_BLACK)

add_text_box(slide, "Think Different.",
    0.5, 1.8, 12.33, 1.3,
    font_size=40, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

add_text_box(slide,
    "Join millions of people who trust Apple design every day.",
    0.5, 3.2, 12.33, 0.8,
    font_size=20, bold=False, color=C_BODY_MUTED, align=PP_ALIGN.CENTER)

pill_button(slide, "Get Started", cx=WI/2, cy=5.5, w=2.2, h=0.45)


# ── Save ─────────────────────────────────────────────────────────────────────
out_path = "/home/user/prd-submissions/apple-design-system-template.pptx"
prs.save(out_path)
print(f"Saved: {out_path}")

import os
size = os.path.getsize(out_path)
print(f"File size: {size:,} bytes ({size/1024:.1f} KB)")
print(f"Slides: {len(prs.slides)}")
