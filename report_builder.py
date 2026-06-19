"""
PPTX report builder — "Midnight Executive" themed deck.

Rewritten from scratch. The original version:
  - dumped raw (often long) text straight into placeholder text boxes with
    no explicit font size, so the layout's default font massively
    overflowed the slide for anything more than a couple of sentences —
    this is the "font too big, doesn't fit the page" complaint
  - had no tables, no SWOT grid, no stat callouts — every content slide
    was just a title plus one giant paragraph
  - rendered financials and recommendations as duplicates of the same
    executive-summary blob (a bug fixed upstream in agents.py)

This version consumes the *structured* fields agents.py now produces
(dicts/lists, not raw strings) and lays them out as real bullets, tables,
and stat callouts, with every font size fixed at a safe value so nothing
overflows a slide.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# --------------------------------------------------------------------------
# Palette — "Midnight Executive" (navy / ice blue / white), one dominant
# color (navy) with ice-blue and white as supporting/accent tones.
# --------------------------------------------------------------------------
NAVY = RGBColor(0x1E, 0x27, 0x61)
NAVY_DARK = RGBColor(0x12, 0x17, 0x42)
ICE = RGBColor(0xCA, 0xDC, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x1B, 0x1E, 0x2B)        # primary body text on light bg
SLATE = RGBColor(0x5B, 0x63, 0x7A)      # secondary / muted text
LIGHT_BG = RGBColor(0xF6, 0xF8, 0xFD)
LINE = RGBColor(0xE2, 0xE7, 0xF2)
GREEN = RGBColor(0x1E, 0x82, 0x53)
RED = RGBColor(0xB8, 0x39, 0x39)
AMBER = RGBColor(0xA8, 0x73, 0x12)
BLUE = RGBColor(0x2A, 0x5C, 0xC8)

FONT_HEAD = "Cambria"
FONT_BODY = "Calibri"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.6)


# ==========================================================================
# Low level helpers
# ==========================================================================

def _blank_slide(prs, bg=WHITE):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = bg
    return slide


def _no_autofit(tf):
    tf.word_wrap = True
    try:
        tf.auto_size = MSO_AUTO_SIZE.NONE
    except Exception:
        pass


def _set_margins(tf, l=0.0, t=0.0, r=0.0, b=0.0):
    tf.margin_left = Inches(l)
    tf.margin_right = Inches(r)
    tf.margin_top = Inches(t)
    tf.margin_bottom = Inches(b)


def add_text(
    slide, x, y, w, h, text,
    size=14, color=INK, bold=False, italic=False,
    font=FONT_BODY, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    line_spacing=1.15, letter_spacing=None,
):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    _no_autofit(tf)
    _set_margins(tf)
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font
    run.font.color.rgb = color
    return box


def add_bullets(
    slide, x, y, w, h, items,
    size=13, color=INK, font=FONT_BODY, bold=False,
    marker="—", marker_color=None, space_after=8, line_spacing=1.12,
):
    """Render a list of short strings as left-aligned bullet paragraphs
    with explicit, safe font sizing (never relies on shape autofit)."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    _no_autofit(tf)
    _set_margins(tf)
    marker_color = marker_color or color

    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        r1 = p.add_run()
        r1.text = f"{marker}  "
        r1.font.size = Pt(size)
        r1.font.name = font
        r1.font.color.rgb = marker_color
        r1.font.bold = True
        r2 = p.add_run()
        r2.text = str(item)
        r2.font.size = Pt(size)
        r2.font.name = font
        r2.font.bold = bold
        r2.font.color.rgb = color
    return box


def add_icon_circle(slide, cx, cy, diameter, fill_color, glyph, glyph_size=16, glyph_color=WHITE):
    """A filled circle (no border) with a short glyph/letter centered in it —
    used as the repeating visual motif instead of an accent stripe."""
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx, cy, diameter, diameter)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    shape.shadow.inherit = False
    tf = shape.text_frame
    _set_margins(tf)
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = glyph
    run.font.size = Pt(glyph_size)
    run.font.bold = True
    run.font.name = FONT_BODY
    run.font.color.rgb = glyph_color
    return shape


def add_section_header(slide, x, y, icon_glyph, icon_color, title, subtitle=None, title_color=NAVY):
    """Repeating header motif: small colored icon circle + bold title."""
    d = Inches(0.46)
    add_icon_circle(slide, x, y, d, icon_color, icon_glyph, glyph_size=15)
    add_text(
        slide, x + d + Inches(0.18), y - Inches(0.02), Inches(9.5), Inches(0.5),
        title, size=22, bold=True, color=title_color, font=FONT_HEAD,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    if subtitle:
        add_text(
            slide, x + d + Inches(0.18), y + Inches(0.42), Inches(9.5), Inches(0.35),
            subtitle, size=12, color=SLATE, font=FONT_BODY,
        )


def add_card(slide, x, y, w, h, fill=WHITE, line_color=LINE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    try:
        shape.adjustments[0] = 0.045
    except Exception:
        pass
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line_color
    shape.line.width = Pt(0.75)
    shape.shadow.inherit = False
    return shape


def add_page_footer(slide, page_no, total, label):
    add_text(
        slide, MARGIN, SLIDE_H - Inches(0.45), Inches(6), Inches(0.3),
        label, size=9, color=SLATE, font=FONT_BODY,
    )
    add_text(
        slide, SLIDE_W - Inches(1.4), SLIDE_H - Inches(0.45), Inches(0.8), Inches(0.3),
        f"{page_no:02d} / {total:02d}", size=9, color=SLATE, font=FONT_BODY, align=PP_ALIGN.RIGHT,
    )


def style_table(table, header_bg=NAVY, header_color=WHITE, body_color=INK, font_size=12, header_size=12, font=FONT_BODY):
    # Strip default theme banding for a flatter, more deliberate look.
    tbl = table._tbl
    style_id = tbl.find(qn("a:tblPr"))
    if style_id is not None:
        style_id.attrib.pop("firstRow", None)

    for c, cell in enumerate(table.rows[0].cells):
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_bg
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.LEFT
            for r in p.runs:
                r.font.bold = True
                r.font.size = Pt(header_size)
                r.font.color.rgb = header_color
                r.font.name = font

    for row in list(table.rows)[1:]:
        for cell in row.cells:
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.LEFT
                for r in p.runs:
                    r.font.size = Pt(font_size)
                    r.font.color.rgb = body_color
                    r.font.name = font


def fill_table(table, headers, rows):
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = str(h)
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = str(val)


# ==========================================================================
# Slide builders
# ==========================================================================

def _slide_title(prs, state, page_no, total):
    slide = _blank_slide(prs, bg=NAVY)

    # soft secondary glow circle as the repeating motif, top-right
    glow = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.2), Inches(-1.6), Inches(4.6), Inches(4.6))
    glow.fill.solid()
    glow.fill.fore_color.rgb = NAVY_DARK
    glow.line.fill.background()
    glow.shadow.inherit = False

    add_text(
        slide, MARGIN, Inches(0.7), Inches(6), Inches(0.4),
        "AI CONSULTANT AGENT TEAM", size=12, bold=True, color=ICE, font=FONT_BODY,
    )

    industry = state.get("industry") or "Target Industry"
    market = state.get("market") or "Target Market"

    add_text(
        slide, MARGIN, Inches(2.5), Inches(11.5), Inches(1.6),
        "Market Entry Strategy", size=44, bold=True, color=WHITE, font=FONT_HEAD,
    )
    add_text(
        slide, MARGIN, Inches(3.55), Inches(11.5), Inches(0.7),
        f"{industry}  ·  {market}", size=20, color=ICE, font=FONT_BODY,
    )

    add_text(
        slide, MARGIN, SLIDE_H - Inches(0.9), Inches(8), Inches(0.4),
        "Prepared by the Market Intelligence, Strategy, and Executive Advisory agents",
        size=11, color=ICE, font=FONT_BODY, italic=True,
    )
    return slide


def _slide_executive_summary(prs, state, page_no, total):
    slide = _blank_slide(prs, bg=WHITE)
    add_section_header(slide, MARGIN, Inches(0.55), "S", NAVY, "Executive Summary",
                        "The opportunity, the strategy, and the call to action")

    card = add_card(slide, MARGIN, Inches(1.6), SLIDE_W - 2 * MARGIN, Inches(3.3), fill=LIGHT_BG, line_color=LINE)

    summary = state.get("executive_summary") or (
        "Executive summary not available for this run."
    )
    add_text(
        slide, MARGIN + Inches(0.4), Inches(1.95), SLIDE_W - 2 * MARGIN - Inches(0.8), Inches(2.7),
        summary, size=16, color=INK, font=FONT_BODY, line_spacing=1.35,
    )

    # three quick stat chips pulled from financials, repeating the icon motif
    fin = state.get("financials") or {}
    chips = [
        ("Initial Investment", fin.get("initial_investment") or "—", BLUE),
        ("Year 1 Revenue", fin.get("year1_revenue") or "—", GREEN),
        ("Breakeven", fin.get("breakeven_timeline") or "—", AMBER),
    ]
    chip_w = Inches(3.7)
    gap = Inches(0.35)
    start_x = MARGIN
    y = Inches(5.25)
    for i, (label, value, color) in enumerate(chips):
        x = start_x + i * (chip_w + gap)
        add_card(slide, x, y, chip_w, Inches(1.5), fill=WHITE, line_color=LINE)
        add_text(slide, x + Inches(0.3), y + Inches(0.18), chip_w - Inches(0.6), Inches(0.4),
                  label.upper(), size=10, bold=True, color=SLATE, font=FONT_BODY)
        add_text(slide, x + Inches(0.3), y + Inches(0.55), chip_w - Inches(0.6), Inches(0.8),
                  value, size=22, bold=True, color=color, font=FONT_HEAD)

    add_page_footer(slide, page_no, total, "Executive Summary")
    return slide


def _slide_market_intelligence(prs, state, page_no, total):
    slide = _blank_slide(prs, bg=WHITE)
    ma = state.get("market_analysis") or {}
    add_section_header(slide, MARGIN, Inches(0.55), "M", BLUE, "Market Intelligence",
                        ma.get("market_size") or "Market sizing and competitive landscape")

    left_w = Inches(6.5)
    right_x = MARGIN + left_w + Inches(0.4)
    right_w = SLIDE_W - right_x - MARGIN

    # Left: overview + growth trends
    add_card(slide, MARGIN, Inches(1.55), left_w, Inches(5.2), fill=LIGHT_BG, line_color=LINE)
    add_text(slide, MARGIN + Inches(0.35), Inches(1.8), left_w - Inches(0.7), Inches(0.3),
              "OVERVIEW", size=11, bold=True, color=SLATE, font=FONT_BODY)
    add_text(slide, MARGIN + Inches(0.35), Inches(2.15), left_w - Inches(0.7), Inches(1.1),
              ma.get("overview") or "Not available.", size=13.5, color=INK, line_spacing=1.3)

    add_text(slide, MARGIN + Inches(0.35), Inches(3.35), left_w - Inches(0.7), Inches(0.3),
              "GROWTH TRENDS", size=11, bold=True, color=SLATE, font=FONT_BODY)
    trends = ma.get("growth_trends") or ["Not available"]
    add_bullets(slide, MARGIN + Inches(0.35), Inches(3.7), left_w - Inches(0.7), Inches(1.9),
                trends, size=13.5, marker="↗", marker_color=GREEN, space_after=10)

    # Right: competitors table + opportunities/risks
    add_text(slide, right_x, Inches(1.8), right_w, Inches(0.3),
              "TOP COMPETITORS", size=11, bold=True, color=SLATE, font=FONT_BODY)
    competitors = ma.get("competitors") or []
    rows = [(c.get("name", "—"), c.get("note", "")) for c in competitors[:3]] or [("—", "Not available")]
    table_h = Inches(0.5 + 0.55 * len(rows))
    gframe = slide.shapes.add_table(len(rows) + 1, 2, right_x, Inches(2.15), right_w, table_h)
    table = gframe.table
    table.columns[0].width = Inches(1.7)
    table.columns[1].width = right_w - Inches(1.7)
    fill_table(table, ["Company", "Why it matters"], rows)
    style_table(table, font_size=11, header_size=11)

    risk_y = Inches(2.2) + table_h + Inches(0.3)
    col_w = (right_w - Inches(0.3)) / 2

    add_text(slide, right_x, risk_y, col_w, Inches(0.3),
              "OPPORTUNITIES", size=11, bold=True, color=SLATE, font=FONT_BODY)
    opps = ma.get("opportunities") or ["Not available"]
    add_bullets(slide, right_x, risk_y + Inches(0.32), col_w, Inches(1.6),
                opps, size=11.5, marker="+", marker_color=GREEN, space_after=8)

    risk_x = right_x + col_w + Inches(0.3)
    add_text(slide, risk_x, risk_y, col_w, Inches(0.3),
              "RISKS", size=11, bold=True, color=SLATE, font=FONT_BODY)
    risks = ma.get("risks") or ["Not available"]
    add_bullets(slide, risk_x, risk_y + Inches(0.32), col_w, Inches(1.6),
                risks, size=11.5, marker="!", marker_color=RED, space_after=8)

    add_page_footer(slide, page_no, total, "Market Intelligence")
    return slide


def _slide_swot(prs, state, page_no, total):
    slide = _blank_slide(prs, bg=WHITE)
    add_section_header(slide, MARGIN, Inches(0.55), "T", NAVY, "Strategy — SWOT Analysis",
                        "Internal strengths/weaknesses and external opportunities/threats")

    swot = (state.get("strategy") or {}).get("swot") or {}
    quadrants = [
        ("STRENGTHS", swot.get("strengths") or ["Not available"], GREEN, RGBColor(0xEC, 0xF7, 0xF0)),
        ("WEAKNESSES", swot.get("weaknesses") or ["Not available"], RED, RGBColor(0xFB, 0xEC, 0xEC)),
        ("OPPORTUNITIES", swot.get("opportunities") or ["Not available"], BLUE, RGBColor(0xEA, 0xF1, 0xFC)),
        ("THREATS", swot.get("threats") or ["Not available"], AMBER, RGBColor(0xFC, 0xF3, 0xE3)),
    ]

    qw = Inches(5.95)
    qh = Inches(2.5)
    gap = Inches(0.3)
    x0 = MARGIN
    y0 = Inches(1.6)
    positions = [
        (x0, y0), (x0 + qw + gap, y0),
        (x0, y0 + qh + gap), (x0 + qw + gap, y0 + qh + gap),
    ]
    for (label, items, color, bg), (x, y) in zip(quadrants, positions):
        add_card(slide, x, y, qw, qh, fill=bg, line_color=color)
        add_text(slide, x + Inches(0.3), y + Inches(0.2), qw - Inches(0.6), Inches(0.35),
                  label, size=13, bold=True, color=color, font=FONT_BODY)
        add_bullets(slide, x + Inches(0.3), y + Inches(0.65), qw - Inches(0.6), qh - Inches(0.85),
                    items[:3], size=12, marker="•", marker_color=color, space_after=6)

    return slide


def _slide_strategy(prs, state, page_no, total):
    slide = _blank_slide(prs, bg=WHITE)
    add_section_header(slide, MARGIN, Inches(0.55), "E", NAVY, "Market Entry Strategy",
                        "Actionable steps to enter and scale in this market")

    entry = (state.get("strategy") or {}).get("entry_strategy") or ["Not available"]
    y = Inches(1.7)
    row_h = Inches(1.05)
    gap = Inches(0.25)
    for i, step in enumerate(entry[:4]):
        card = add_card(slide, MARGIN, y, SLIDE_W - 2 * MARGIN, row_h, fill=LIGHT_BG, line_color=LINE)
        add_icon_circle(slide, MARGIN + Inches(0.3), y + Inches(0.25), Inches(0.55), NAVY, str(i + 1), glyph_size=18)
        add_text(slide, MARGIN + Inches(1.1), y, SLIDE_W - 2 * MARGIN - Inches(1.5), row_h,
                  step, size=15, color=INK, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.2)
        y += row_h + gap

    return slide


def _slide_financials(prs, state, page_no, total):
    slide = _blank_slide(prs, bg=WHITE)
    add_section_header(slide, MARGIN, Inches(0.55), "$", GREEN, "Financial Estimate",
                        "Investment, revenue outlook, and path to breakeven")

    fin = state.get("financials") or {}

    chips = [
        ("Initial Investment", fin.get("initial_investment") or "—", BLUE),
        ("Year 1 Revenue", fin.get("year1_revenue") or "—", GREEN),
        ("Breakeven Timeline", fin.get("breakeven_timeline") or "—", AMBER),
    ]
    chip_w = Inches(3.7)
    gap = Inches(0.35)
    y = Inches(1.65)
    for i, (label, value, color) in enumerate(chips):
        x = MARGIN + i * (chip_w + gap)
        add_card(slide, x, y, chip_w, Inches(1.5), fill=LIGHT_BG, line_color=LINE)
        add_text(slide, x + Inches(0.3), y + Inches(0.18), chip_w - Inches(0.6), Inches(0.4),
                  label.upper(), size=10, bold=True, color=SLATE)
        add_text(slide, x + Inches(0.3), y + Inches(0.55), chip_w - Inches(0.6), Inches(0.8),
                  value, size=24, bold=True, color=color, font=FONT_HEAD)

    # 3-year revenue projection table + simple bar visual
    projection = fin.get("projection") or []
    table_y = Inches(3.5)
    add_text(slide, MARGIN, Inches(3.2), Inches(6), Inches(0.3),
              "REVENUE PROJECTION", size=11, bold=True, color=SLATE)

    if projection:
        rows = [(p.get("year", "—"), p.get("revenue", "—")) for p in projection[:3]]
        gframe = slide.shapes.add_table(len(rows) + 1, 2, MARGIN, table_y, Inches(4.6), Inches(0.5 + 0.55 * len(rows)))
        table = gframe.table
        table.columns[0].width = Inches(2.0)
        table.columns[1].width = Inches(2.6)
        fill_table(table, ["Period", "Revenue"], rows)
        style_table(table, font_size=13, header_size=12)

        # simple horizontal bar chart next to the table, parsed loosely from
        # the dollar figures so it degrades gracefully if parsing fails
        import re as _re

        def _to_number(s):
            digits = _re.sub(r"[^0-9.]", "", str(s))
            try:
                return float(digits) if digits else 0.0
            except ValueError:
                return 0.0

        values = [_to_number(p.get("revenue", "0")) for p in projection[:3]]
        max_v = max(values) or 1.0
        bx_in = 6.0          # left edge of bar track, inches
        track_w_in = 5.4      # max bar width, inches
        min_bar_in = 0.5       # minimum visible bar width, inches
        by = table_y + Inches(0.05)
        bar_h = Inches(0.5)
        bar_gap = Inches(0.25)
        for i, (p, v) in enumerate(zip(projection[:3], values)):
            frac = (v / max_v) if max_v else 0.0
            bar_w_in = min_bar_in + (track_w_in - min_bar_in) * frac
            bar = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(bx_in), by + i * (bar_h + bar_gap),
                Inches(bar_w_in), bar_h,
            )
            bar.fill.solid()
            bar.fill.fore_color.rgb = [BLUE, NAVY, GREEN][i % 3]
            bar.line.fill.background()
            bar.shadow.inherit = False
            try:
                bar.adjustments[0] = 0.25
            except Exception:
                pass
            add_text(slide, Inches(bx_in + bar_w_in + 0.15), by + i * (bar_h + bar_gap), Inches(1.4), bar_h,
                      p.get("revenue", ""), size=12, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
    else:
        add_text(slide, MARGIN, table_y, Inches(8), Inches(0.5), "Not available.", size=13, color=SLATE)

    return slide


def _slide_recommendations(prs, state, page_no, total):
    slide = _blank_slide(prs, bg=NAVY)
    add_text(slide, MARGIN, Inches(0.7), Inches(10), Inches(0.6),
              "Recommendations", size=30, bold=True, color=WHITE, font=FONT_HEAD)
    add_text(slide, MARGIN, Inches(1.3), Inches(10), Inches(0.4),
              "What the partner team advises doing next", size=13, color=ICE)

    recs = state.get("recommendations") or ["Not available"]
    y = Inches(2.1)
    row_h = Inches(1.05)
    gap = Inches(0.22)
    for i, rec in enumerate(recs[:4]):
        add_icon_circle(slide, MARGIN, y + Inches(0.2), Inches(0.6), ICE, str(i + 1), glyph_size=20, glyph_color=NAVY)
        add_text(slide, MARGIN + Inches(0.85), y, SLIDE_W - 2 * MARGIN - Inches(0.85), row_h,
                  rec, size=16, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.2)
        y += row_h + gap

    return slide


def _slide_closing(prs, state, page_no, total):
    slide = _blank_slide(prs, bg=NAVY)
    glow = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-1.4), Inches(4.2), Inches(4.6), Inches(4.6))
    glow.fill.solid()
    glow.fill.fore_color.rgb = NAVY_DARK
    glow.line.fill.background()
    glow.shadow.inherit = False

    add_text(slide, MARGIN, Inches(3.0), Inches(11.5), Inches(0.9),
              "Thank You", size=40, bold=True, color=WHITE, font=FONT_HEAD)
    add_text(slide, MARGIN, Inches(3.85), Inches(11.5), Inches(0.5),
              "Generated by the AI Consultant Agent Team", size=14, color=ICE)
    return slide


# ==========================================================================
# Entry point
# ==========================================================================

def build_presentation(state: dict, output_path: str = "market_entry_report.pptx") -> str:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    builders = [
        _slide_title,
        _slide_executive_summary,
        _slide_market_intelligence,
        _slide_swot,
        _slide_strategy,
        _slide_financials,
        _slide_recommendations,
        _slide_closing,
    ]
    total = len(builders)
    for i, builder in enumerate(builders, start=1):
        builder(prs, state, i, total)

    prs.save(output_path)
    return output_path


def build_markdown_report(state: dict) -> str:
    """Plain-text/Markdown export of the same report — quick to read or
    paste elsewhere, complementary to the PPTX download."""
    industry = state.get("industry", "")
    market = state.get("market", "")
    ma = state.get("market_analysis") or {}
    strat = state.get("strategy") or {}
    swot = strat.get("swot") or {}
    fin = state.get("financials") or {}

    def bullets(items):
        return "\n".join(f"- {i}" for i in (items or [])) or "- Not available"

    lines = [
        f"# Market Entry Strategy — {industry} / {market}",
        "",
        "## Executive Summary",
        state.get("executive_summary") or "Not available.",
        "",
        "## Market Intelligence",
        f"**Overview:** {ma.get('overview', '')}",
        f"**Market size:** {ma.get('market_size', '')}",
        "",
        "**Growth trends:**",
        bullets(ma.get("growth_trends")),
        "",
        "**Top competitors:**",
        "\n".join(
            f"- {c.get('name', '—')} — {c.get('note', '')}"
            for c in (ma.get("competitors") or [])
        ) or "- Not available",
        "",
        "**Opportunities:**",
        bullets(ma.get("opportunities")),
        "",
        "**Risks:**",
        bullets(ma.get("risks")),
        "",
        "## SWOT Analysis",
        "**Strengths:**", bullets(swot.get("strengths")), "",
        "**Weaknesses:**", bullets(swot.get("weaknesses")), "",
        "**Opportunities:**", bullets(swot.get("opportunities")), "",
        "**Threats:**", bullets(swot.get("threats")), "",
        "## Market Entry Strategy",
        bullets(strat.get("entry_strategy")),
        "",
        "## Financial Estimate",
        f"- Initial investment: {fin.get('initial_investment', '—')}",
        f"- Year 1 revenue: {fin.get('year1_revenue', '—')}",
        f"- Breakeven timeline: {fin.get('breakeven_timeline', '—')}",
        "",
        "**Revenue projection:**",
        "\n".join(
            f"- {p.get('year', '—')}: {p.get('revenue', '—')}"
            for p in (fin.get("projection") or [])
        ) or "- Not available",
        "",
        "## Recommendations",
        bullets(state.get("recommendations")),
    ]
    return "\n".join(lines)