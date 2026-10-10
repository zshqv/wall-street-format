"""
Generate a sample Wall Street-formatted PowerPoint pitchbook for Titan Industries.
Follows institutional IB/PE slide design standards.
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

# ---------------------------------------------------------------------------
# Color Palette
# ---------------------------------------------------------------------------
NAVY = RGBColor(0x0A, 0x25, 0x40)
DARK_NAVY = RGBColor(0x0B, 0x23, 0x41)
SLATE = RGBColor(0x4A, 0x55, 0x68)
OFF_WHITE = RGBColor(0xF7, 0xFA, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BURGUNDY = RGBColor(0x8B, 0x00, 0x00)
GOLD = RGBColor(0xC5, 0xA0, 0x59)
LIGHT_GRAY = RGBColor(0xE2, 0xE8, 0xF0)
MID_GRAY = RGBColor(0x71, 0x80, 0x96)
BLACK = RGBColor(0x00, 0x00, 0x00)

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)


def _set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_textbox(slide, left, top, width, height, text, font_size=12,
                 color=BLACK, bold=False, alignment=PP_ALIGN.LEFT, font_name="Arial"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def _add_rect(slide, left, top, width, height, fill_color, text="",
              font_size=11, font_color=WHITE, bold=False):
    shape = slide.shapes.add_shape(1, left, top, width, height)  # 1 = rectangle
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    if text:
        tf = shape.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].text = text
        tf.paragraphs[0].font.size = Pt(font_size)
        tf.paragraphs[0].font.color.rgb = font_color
        tf.paragraphs[0].font.bold = bold
        tf.paragraphs[0].font.name = "Arial"
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        shape.text_frame.paragraphs[0].space_before = Pt(0)
    return shape


def _add_tracker(slide, text):
    _add_textbox(slide, Inches(0.7), Inches(0.35), Inches(10), Inches(0.35),
                 text, font_size=9, color=MID_GRAY, bold=False)


def _add_action_title(slide, text):
    _add_textbox(slide, Inches(0.7), Inches(0.65), Inches(11.5), Inches(0.7),
                 text, font_size=24, color=NAVY, bold=True)


def _add_divider(slide, top):
    shape = slide.shapes.add_shape(1, Inches(0.7), top, Inches(11.9), Pt(2))
    shape.fill.solid()
    shape.fill.fore_color.rgb = GOLD
    shape.line.fill.background()


# ===================================================================
# SLIDE 1: COVER
# ===================================================================
def slide_cover(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_slide_bg(slide, NAVY)

    _add_textbox(slide, Inches(0.7), Inches(0.5), Inches(4), Inches(0.3),
                 "CONFIDENTIAL", font_size=10, color=GOLD, bold=True)

    shape = slide.shapes.add_shape(1, Inches(0.7), Inches(2.2), Inches(11.9), Pt(3))
    shape.fill.solid()
    shape.fill.fore_color.rgb = GOLD
    shape.line.fill.background()

    _add_textbox(slide, Inches(0.7), Inches(2.6), Inches(11), Inches(1.2),
                 "Titan Industries, Inc.", font_size=40, color=WHITE, bold=True)

    _add_textbox(slide, Inches(0.7), Inches(3.8), Inches(11), Inches(0.6),
                 "Confidential Information Memorandum", font_size=20, color=LIGHT_GRAY)

    _add_textbox(slide, Inches(0.7), Inches(5.2), Inches(11), Inches(0.4),
                 "Project Titan  |  October 2026", font_size=14, color=MID_GRAY)

    _add_textbox(slide, Inches(0.7), Inches(6.5), Inches(11), Inches(0.4),
                 "Prepared by Titan Capital Advisory", font_size=11, color=MID_GRAY)


# ===================================================================
# SLIDE 2: EXECUTIVE SUMMARY
# ===================================================================
def slide_executive_summary(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_slide_bg(slide, WHITE)

    _add_tracker(slide, "Project Titan  |  Executive Summary")
    _add_action_title(slide, "Titan Industries Is a Market Leader Positioned for Accelerated Growth")
    _add_divider(slide, Inches(1.35))

    cards = [
        ("Company Overview",
         "Titan Industries is a diversified industrial technology company with $2.4B in "
         "FY2024E revenue, serving enterprise clients across manufacturing, logistics, and "
         "energy sectors with mission-critical software and hardware solutions."),
        ("Growth Trajectory",
         "Revenue has grown at a 15% CAGR over the past 3 years, driven by enterprise SaaS "
         "expansion and strategic M&A. EBITDA margins are expanding from 22% to 28% through "
         "operating leverage and product mix shift."),
        ("Investment Thesis",
         "The company represents a compelling platform investment opportunity with multiple "
         "organic and inorganic growth levers, strong FCF generation ($279M FY2024E), and a "
         "clear path to $1B+ EBITDA within 5 years."),
    ]

    for i, (title, body) in enumerate(cards):
        left = Inches(0.7 + i * 4.1)
        _add_rect(slide, left, Inches(1.7), Inches(3.8), Inches(0.5),
                  NAVY, title, font_size=13, font_color=WHITE, bold=True)
        _add_textbox(slide, left + Inches(0.15), Inches(2.3), Inches(3.5), Inches(3.5),
                     body, font_size=11, color=SLATE)


# ===================================================================
# SLIDE 3: FINANCIAL OVERVIEW
# ===================================================================
def slide_financial_overview(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_slide_bg(slide, WHITE)

    _add_tracker(slide, "Project Titan  |  Financial Overview")
    _add_action_title(slide,
        "Revenue Growing at 15% CAGR with EBITDA Margins Expanding 600bps Over Forecast Period")
    _add_divider(slide, Inches(1.35))

    metrics = [
        ("$2.4B", "FY2024E\nRevenue", NAVY),
        ("26.0%", "FY2024E\nEBITDA Margin", NAVY),
        ("$279M", "FY2024E\nFree Cash Flow", NAVY),
        ("15.0%", "Revenue\nCAGR ('22-'24)", BURGUNDY),
    ]

    for i, (val, label, color) in enumerate(metrics):
        left = Inches(0.7 + i * 3.15)
        _add_rect(slide, left, Inches(1.7), Inches(2.8), Inches(1.5),
                  OFF_WHITE, "", font_size=11)
        _add_textbox(slide, left + Inches(0.2), Inches(1.8), Inches(2.4), Inches(0.7),
                     val, font_size=32, color=color, bold=True, alignment=PP_ALIGN.CENTER)
        _add_textbox(slide, left + Inches(0.2), Inches(2.5), Inches(2.4), Inches(0.6),
                     label, font_size=11, color=SLATE, alignment=PP_ALIGN.CENTER)

    headers = ["($M)", "FY 2022A", "FY 2023A", "FY 2024E", "FY 2025E", "FY 2026E"]
    rows = [
        ["Revenue", "$1,850", "$2,120", "$2,438", "$2,804", "$3,164"],
        ["Gross Profit", "$925", "$1,081", "$1,268", "$1,458", "$1,645"],
        ["EBITDA", "$407", "$508", "$634", "$757", "$886"],
        ["Net Income", "$170", "$241", "$328", "$411", "$500"],
        ["Free Cash Flow", "$123", "$198", "$279", "$353", "$437"],
    ]

    table_top = Inches(3.5)
    col_w = [Inches(2.2)] + [Inches(1.9)] * 5
    row_h = Inches(0.4)

    tbl = slide.shapes.add_table(len(rows) + 1, 6,
                                  Inches(0.7), table_top,
                                  Inches(11.7), row_h * (len(rows) + 1)).table

    for c, w in enumerate(col_w):
        tbl.columns[c].width = w

    for c, h in enumerate(headers):
        cell = tbl.cell(0, c)
        cell.text = h
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.font.name = "Arial"
            p.alignment = PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY

    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = tbl.cell(r + 1, c)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(10)
                p.font.color.rgb = BLACK
                p.font.name = "Arial"
                p.font.bold = (r == len(rows) - 1)
                p.alignment = PP_ALIGN.RIGHT if c > 0 else PP_ALIGN.LEFT
            cell.fill.solid()
            cell.fill.fore_color.rgb = OFF_WHITE if r % 2 == 0 else WHITE


# ===================================================================
# SLIDE 4: VALUATION SUMMARY
# ===================================================================
def slide_valuation(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_slide_bg(slide, WHITE)

    _add_tracker(slide, "Project Titan  |  Valuation Summary")
    _add_action_title(slide,
        "DCF Analysis Implies $24.69 Per Share; LBO Supports 2.4x MOIC at 33.8% IRR")
    _add_divider(slide, Inches(1.35))

    for i, (title, items) in enumerate([
        ("DCF Valuation", [
            "WACC: 9.0%",
            "Terminal Growth: 2.5%",
            "Enterprise Value: $5,483M",
            "Equity Value: $6,173M",
            "Implied Share Price: $24.69",
        ]),
        ("LBO Returns Analysis", [
            "Entry Multiple: 10.0x EBITDA",
            "Leverage: 6.0x at Entry",
            "Exit Multiple: 10.0x EBITDA",
            "MOIC: 2.4x",
            "IRR: 33.8%",
        ]),
    ]):
        left = Inches(0.7 + i * 6.2)
        _add_rect(slide, left, Inches(1.7), Inches(5.8), Inches(0.5),
                  NAVY, title, font_size=14, font_color=WHITE, bold=True)
        for j, item in enumerate(items):
            parts = item.split(": ")
            _add_textbox(slide, left + Inches(0.3), Inches(2.4 + j * 0.55),
                         Inches(5.2), Inches(0.45),
                         item, font_size=13, color=SLATE)

    _add_rect(slide, Inches(0.7), Inches(5.8), Inches(11.9), Inches(1.1),
              NAVY, "", font_size=11)
    _add_textbox(slide, Inches(1.0), Inches(5.9), Inches(11.3), Inches(0.9),
                 "Recommendation: Titan represents an attractive acquisition target at current "
                 "valuations. The combination of secular tailwinds, margin expansion trajectory, and "
                 "strong FCF generation supports a premium valuation. We recommend proceeding to "
                 "Phase II due diligence.",
                 font_size=12, color=WHITE, bold=False)


# ===================================================================
# SLIDE 5: APPENDIX / KEY RISKS
# ===================================================================
def slide_risks(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_slide_bg(slide, WHITE)

    _add_tracker(slide, "Project Titan  |  Risk Factors & Appendix")
    _add_action_title(slide,
        "Key Risks Are Manageable with Appropriate Structuring and Diligence Focus Areas")
    _add_divider(slide, Inches(1.35))

    risks = [
        ("Customer Concentration",
         "Top 10 customers represent ~35% of revenue. Mitigation: diversified end-market "
         "exposure and long-term contracts with 95%+ renewal rates."),
        ("Integration Risk",
         "Prior acquisitions require continued integration effort. Mitigation: dedicated PMO "
         "with track record of achieving >90% of targeted synergies within 18 months."),
        ("Macro Sensitivity",
         "Industrial end-markets correlate with GDP cycles. Mitigation: 60% recurring revenue "
         "base with high switching costs provides downside protection."),
        ("Technology Disruption",
         "Emerging competitors in adjacent software verticals. Mitigation: $195M+ annual R&D "
         "investment and 50+ patents protect competitive moat."),
    ]

    for i, (title, body) in enumerate(risks):
        row = i // 2
        col = i % 2
        left = Inches(0.7 + col * 6.2)
        top = Inches(1.7 + row * 2.6)
        _add_rect(slide, left, top, Inches(5.8), Inches(0.45),
                  BURGUNDY, title, font_size=12, font_color=WHITE, bold=True)
        _add_textbox(slide, left + Inches(0.15), top + Inches(0.55),
                     Inches(5.5), Inches(1.8),
                     body, font_size=11, color=SLATE)


# ===================================================================
# MAIN
# ===================================================================
def main():
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    slide_cover(prs)
    slide_executive_summary(prs)
    slide_financial_overview(prs)
    slide_valuation(prs)
    slide_risks(prs)

    out = Path(__file__).parent.parent / "assets" / "Titan_Industries_Pitchbook.pptx"
    prs.save(str(out))
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
