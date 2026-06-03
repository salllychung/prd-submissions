from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import ChartData
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pptx.oxml
from lxml import etree
import copy

# ── Colors ────────────────────────────────────────────────────────────────────
PRIMARY_BLUE   = RGBColor(0x00, 0x66, 0xCC)
FOCUS_BLUE     = RGBColor(0x00, 0x71, 0xE3)
SKY_BLUE       = RGBColor(0x29, 0x97, 0xFF)
CANVAS_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
PARCHMENT      = RGBColor(0xF5, 0xF5, 0xF7)
NEAR_BLACK     = RGBColor(0x1D, 0x1D, 0x1F)
DARK_TILE1     = RGBColor(0x27, 0x27, 0x29)
DARK_TILE2     = RGBColor(0x2A, 0x2A, 0x2C)
PURE_BLACK     = RGBColor(0x00, 0x00, 0x00)
BODY_MUTED     = RGBColor(0xCC, 0xCC, 0xCC)
HAIRLINE       = RGBColor(0xE0, 0xE0, 0xE0)
INK_MUTED      = RGBColor(0x33, 0x33, 0x33)

FONT = "Calibri"

# ── Helpers ───────────────────────────────────────────────────────────────────
def set_bg(slide, color: RGBColor):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_text_box(slide, text, left, top, width, height,
                 font_size, bold=False, color=NEAR_BLACK,
                 align=PP_ALIGN.LEFT, wrap=True, font=FONT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox

def add_rect(slide, left, top, width, height, fill_color, line_color=None, line_width=Pt(0.5)):
    from pptx.util import Pt as _Pt
    shape = slide.shapes.add_shape(
        pptx.enum.shapes.MSO_SHAPE_TYPE.AUTO,  # placeholder; overridden below
        left, top, width, height
    )
    # use add_shape with freeform? No – simpler: just add_textbox then skip.
    # Actually use slide.shapes.add_shape with a rectangle enum value
    return shape

def add_rectangle(slide, left, top, width, height, fill_color, line_color=None):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.util import Pt
    # MSO_SHAPE_TYPE is not for add_shape; use integer 1 for RECTANGLE
    shape = slide.shapes.add_shape(1, left, top, width, height)  # 1 = RECTANGLE
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(0.5)
    else:
        shape.line.fill.background()  # no line
    return shape

def add_rounded_rect(slide, left, top, width, height, fill_color, radius_pct=50, line_color=None):
    """Add rounded rectangle. radius_pct: 0-100, where 50 = pill for small heights."""
    shape = slide.shapes.add_shape(5, left, top, width, height)  # 5 = ROUNDED_RECTANGLE
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    # Set corner radius via XML adjustment
    adj = shape.adjustments
    if len(adj) > 0:
        adj[0] = radius_pct / 100.0 * 50000  # EMU units for adjustment
    return shape

def pill_button(slide, label, cx, cy, w=Inches(1.5), h=Inches(0.4),
                fill=PRIMARY_BLUE, text_color=CANVAS_WHITE, font_size=10):
    """Centered pill button at (cx, cy) center point."""
    left = cx - w / 2
    top  = cy - h / 2
    shape = slide.shapes.add_shape(5, left, top, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    # Max corner radius
    try:
        shape.adjustments[0] = 50000
    except Exception:
        pass
    tf = shape.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = label
    run.font.name = FONT
    run.font.size = Pt(font_size)
    run.font.bold = False
    run.font.color.rgb = text_color
    return shape

def chart_font(chart, font_name=FONT):
    """Set chart-wide font to font_name."""
    try:
        txPr = chart._element.find('.//' + qn('c:txPr'))
        if txPr is None:
            return
        for latin in txPr.iter(qn('a:latin')):
            latin.set('typeface', font_name)
    except Exception:
        pass

# ── Presentation setup ────────────────────────────────────────────────────────
prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

W = prs.slide_width
H = prs.slide_height

blank_layout = prs.slide_layouts[6]  # completely blank

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 1 – Title Slide Light
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, CANVAS_WHITE)

add_text_box(slide, "Apple Design System",
             Inches(1), Inches(2.2), Inches(11.33), Inches(1.2),
             40, bold=True, color=NEAR_BLACK, align=PP_ALIGN.CENTER)

add_text_box(slide, "A comprehensive presentation template built on Apple's visual language",
             Inches(1.5), Inches(3.5), Inches(10.33), Inches(0.8),
             20, bold=False, color=NEAR_BLACK, align=PP_ALIGN.CENTER)

# Two pill CTAs
pill_button(slide, "Learn More",
            cx=W/2 - Inches(1.1), cy=Inches(5.8),
            w=Inches(1.8), h=Inches(0.45))
pill_button(slide, "Get Started",
            cx=W/2 + Inches(1.1), cy=Inches(5.8),
            w=Inches(1.8), h=Inches(0.45))

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 2 – Title Slide Dark
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, DARK_TILE1)

add_text_box(slide, "Apple Design System",
             Inches(1), Inches(2.2), Inches(11.33), Inches(1.2),
             40, bold=True, color=CANVAS_WHITE, align=PP_ALIGN.CENTER)

add_text_box(slide, "A comprehensive presentation template built on Apple's visual language",
             Inches(1.5), Inches(3.5), Inches(10.33), Inches(0.8),
             20, bold=False, color=BODY_MUTED, align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 3 – Section Divider
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, DARK_TILE1)

add_text_box(slide, "01",
             Inches(1), Inches(1.5), Inches(4), Inches(1.2),
             56, bold=True, color=PRIMARY_BLUE, align=PP_ALIGN.LEFT)

add_text_box(slide, "Introduction to Design",
             Inches(1), Inches(2.9), Inches(8), Inches(1.0),
             28, bold=True, color=CANVAS_WHITE, align=PP_ALIGN.LEFT)

add_text_box(slide, "Explore the principles that define our visual language",
             Inches(1), Inches(4.1), Inches(8), Inches(0.7),
             15, bold=False, color=BODY_MUTED, align=PP_ALIGN.LEFT)

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 4 – Content + Body
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, PARCHMENT)

# Left 60%
add_text_box(slide, "Design Principles",
             Inches(0.6), Inches(1.0), Inches(7.4), Inches(0.7),
             24, bold=True, color=NEAR_BLACK)

body_text = (
    "Apple's design language is rooted in clarity, deference, and depth. "
    "Every element is intentional, every interaction considered.\n\n"
    "Our visual system provides consistency across all touchpoints, ensuring "
    "that users always feel at home within our ecosystem."
)
add_text_box(slide, body_text,
             Inches(0.6), Inches(1.9), Inches(7.4), Inches(3.8),
             12, bold=False, color=NEAR_BLACK, wrap=True)

# Right 40% – image placeholder rounded rect
img_ph = slide.shapes.add_shape(5,  # ROUNDED_RECTANGLE
    Inches(8.4), Inches(0.9), Inches(4.3), Inches(5.4))
img_ph.fill.solid()
img_ph.fill.fore_color.rgb = HAIRLINE
img_ph.line.color.rgb = HAIRLINE
img_ph.line.width = Pt(0.5)
try:
    img_ph.adjustments[0] = 5000  # ~18pt radius feel
except Exception:
    pass

# Label inside placeholder
add_text_box(slide, "Image Placeholder",
             Inches(8.4), Inches(3.3), Inches(4.3), Inches(0.5),
             12, bold=False, color=BODY_MUTED, align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 5 – Full Bullets
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, CANVAS_WHITE)

add_text_box(slide, "Key Features & Benefits",
             Inches(0.8), Inches(0.7), Inches(11.73), Inches(0.7),
             24, bold=True, color=NEAR_BLACK)

# Horizontal rule using thin rectangle
rule = add_rectangle(slide,
    Inches(0.8), Inches(1.55), Inches(11.73), Pt(1),
    HAIRLINE)

bullets = [
    "Clarity — Remove complexity, let content speak for itself",
    "Deference — UI supports without competing with content",
    "Depth — Visual layers create hierarchy and focus",
    "Consistency — Predictable patterns build confidence",
    "Feedback — Every action has a clear, immediate response",
]

for i, text in enumerate(bullets):
    y = Inches(1.85) + i * Inches(0.95)
    # Blue square bullet
    sq = add_rectangle(slide,
        Inches(0.8), y + Inches(0.08), Inches(0.12), Inches(0.12),
        PRIMARY_BLUE)
    add_text_box(slide, text,
                 Inches(1.1), y, Inches(11.0), Inches(0.7),
                 12, bold=False, color=NEAR_BLACK)

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 6 – Bar Chart
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, PARCHMENT)

add_text_box(slide, "Performance Metrics",
             Inches(0.6), Inches(0.4), Inches(11.73), Inches(0.6),
             24, bold=True, color=NEAR_BLACK)

chart_data = ChartData()
chart_data.categories = ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024']
chart_data.add_series('Product A', (42, 67, 58, 83))
chart_data.add_series('Product B', (28, 45, 71, 62))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(0.6), Inches(1.2), Inches(12.13), Inches(5.7),
    chart_data
)
chart = chart_frame.chart
chart.has_legend = True

# Color the series
series_colors = [PRIMARY_BLUE, SKY_BLUE]
for idx, series in enumerate(chart.series):
    fill = series.format.fill
    fill.solid()
    fill.fore_color.rgb = series_colors[idx]

# Style chart area
# (chart_area not accessible via python-pptx API directly)

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 7 – Line Chart
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, CANVAS_WHITE)

add_text_box(slide, "Growth Trends",
             Inches(0.6), Inches(0.4), Inches(11.73), Inches(0.6),
             24, bold=True, color=NEAR_BLACK)

chart_data = ChartData()
chart_data.categories = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
chart_data.add_series('Revenue', (100, 118, 132, 145, 162, 189))
chart_data.add_series('Users',   (80,  95,  110, 128, 150, 175))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.LINE,
    Inches(0.6), Inches(1.2), Inches(12.13), Inches(5.7),
    chart_data
)
chart = chart_frame.chart
chart.has_legend = True
# (chart_area not accessible via python-pptx API directly)

line_colors = [PRIMARY_BLUE, FOCUS_BLUE]
for idx, series in enumerate(chart.series):
    series.format.line.color.rgb = line_colors[idx]
    series.format.line.width = Pt(2.5)

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 8 – Pie Chart
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, DARK_TILE1)

add_text_box(slide, "Market Distribution",
             Inches(0.6), Inches(0.4), Inches(11.73), Inches(0.6),
             24, bold=True, color=CANVAS_WHITE)

chart_data = ChartData()
chart_data.categories = ['iPhone', 'Mac', 'iPad', 'Wearables', 'Services']
chart_data.add_series('Revenue Share', (38, 14, 9, 11, 28))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.PIE,
    Inches(2.0), Inches(1.0), Inches(9.33), Inches(6.0),
    chart_data
)
chart = chart_frame.chart
chart.has_legend = True
# (chart_area not accessible via python-pptx API directly)

pie_colors = [PRIMARY_BLUE, SKY_BLUE, FOCUS_BLUE, BODY_MUTED, PARCHMENT]
for idx, point in enumerate(chart.series[0].points):
    point.format.fill.solid()
    point.format.fill.fore_color.rgb = pie_colors[idx % len(pie_colors)]

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 9 – Data Table
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, CANVAS_WHITE)

add_text_box(slide, "Quarterly Summary",
             Inches(0.6), Inches(0.4), Inches(11.73), Inches(0.6),
             24, bold=True, color=NEAR_BLACK)

rows, cols = 5, 4
table_shape = slide.shapes.add_table(
    rows, cols,
    Inches(0.6), Inches(1.3), Inches(12.13), Inches(4.8)
)
tbl = table_shape.table

headers = ['Category', 'Q1 2024', 'Q2 2024', 'Q3 2024']
row_data = [
    ['iPhone',     '$51.3B', '$55.1B', '$58.7B'],
    ['Mac',        '$7.2B',  '$8.0B',  '$9.3B' ],
    ['iPad',       '$5.6B',  '$6.1B',  '$6.9B' ],
    ['Wearables',  '$9.8B',  '$10.2B', '$11.1B'],
]

# Column widths
col_widths = [Inches(3.5), Inches(2.9), Inches(2.9), Inches(2.83)]
for ci, w in enumerate(col_widths):
    tbl.columns[ci].width = w

# Header row
for ci, hdr in enumerate(headers):
    cell = tbl.cell(0, ci)
    cell.fill.solid()
    cell.fill.fore_color.rgb = NEAR_BLACK
    tf = cell.text_frame
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = tf.paragraphs[0].add_run()
    run.text = hdr
    run.font.name = FONT
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = CANVAS_WHITE

# Data rows
row_fills = [CANVAS_WHITE, PARCHMENT]
for ri, row in enumerate(row_data):
    bg = row_fills[ri % 2]
    for ci, val in enumerate(row):
        cell = tbl.cell(ri + 1, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
        tf = cell.text_frame
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = tf.paragraphs[0].add_run()
        run.text = val
        run.font.name = FONT
        run.font.size = Pt(12)
        run.font.color.rgb = NEAR_BLACK

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 10 – Quote / Callout
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, DARK_TILE1)

add_text_box(slide, "❝",
             Inches(0.6), Inches(0.7), Inches(3), Inches(1.5),
             72, bold=True, color=PRIMARY_BLUE, align=PP_ALIGN.LEFT)

add_text_box(slide,
    "Design is not just what it looks like and feels like.\nDesign is how it works.",
    Inches(1.0), Inches(2.4), Inches(11.33), Inches(1.8),
    24, bold=False, color=CANVAS_WHITE, align=PP_ALIGN.CENTER)

add_text_box(slide, "— Steve Jobs",
             Inches(1.0), Inches(4.4), Inches(11.33), Inches(0.6),
             14, bold=False, color=BODY_MUTED, align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 11 – Two-Column Comparison
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, CANVAS_WHITE)

add_text_box(slide, "Side-by-Side Comparison",
             Inches(0.6), Inches(0.4), Inches(12.13), Inches(0.65),
             24, bold=True, color=NEAR_BLACK, align=PP_ALIGN.CENTER)

# Left box
left_box = add_rectangle(slide,
    Inches(0.5), Inches(1.3), Inches(5.9), Inches(5.5),
    PARCHMENT)

add_text_box(slide, "Option A",
             Inches(0.7), Inches(1.6), Inches(5.5), Inches(0.55),
             15, bold=True, color=NEAR_BLACK)

left_bullets = [
    "Clear and minimal interface design",
    "Optimized for everyday workflows",
    "Seamless device integration",
]
for i, t in enumerate(left_bullets):
    add_text_box(slide, f"• {t}",
                 Inches(0.7), Inches(2.4) + i * Inches(0.8), Inches(5.4), Inches(0.6),
                 12, color=NEAR_BLACK)

# Right box
right_box = add_rectangle(slide,
    Inches(6.93), Inches(1.3), Inches(5.9), Inches(5.5),
    DARK_TILE1)

add_text_box(slide, "Option B",
             Inches(7.13), Inches(1.6), Inches(5.5), Inches(0.55),
             15, bold=True, color=CANVAS_WHITE)

right_bullets = [
    "Immersive dark mode experience",
    "Engineered for power users",
    "Advanced customization options",
]
for i, t in enumerate(right_bullets):
    add_text_box(slide, f"• {t}",
                 Inches(7.13), Inches(2.4) + i * Inches(0.8), Inches(5.4), Inches(0.6),
                 12, color=CANVAS_WHITE)

# ═══════════════════════════════════════════════════════════════════════════════
# Slide 12 – Closing Slide
# ═══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(blank_layout)
set_bg(slide, PURE_BLACK)

add_text_box(slide, "The Future is Here.",
             Inches(1), Inches(1.8), Inches(11.33), Inches(1.3),
             40, bold=True, color=CANVAS_WHITE, align=PP_ALIGN.CENTER)

add_text_box(slide, "Join us in building the next generation of experiences.",
             Inches(1.5), Inches(3.2), Inches(10.33), Inches(0.8),
             20, bold=False, color=BODY_MUTED, align=PP_ALIGN.CENTER)

pill_button(slide, "Get Started",
            cx=W/2, cy=Inches(5.5),
            w=Inches(2.2), h=Inches(0.5),
            fill=PRIMARY_BLUE, text_color=CANVAS_WHITE, font_size=12)

# ── Save ──────────────────────────────────────────────────────────────────────
output_path = "/home/user/prd-submissions/apple-design-system-template.pptx"
prs.save(output_path)
print(f"Saved: {output_path}")

import os
size = os.path.getsize(output_path)
print(f"File size: {size:,} bytes ({size/1024:.1f} KB)")
print(f"Slides: {len(prs.slides)}")
