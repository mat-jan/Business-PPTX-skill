"""
apply_theme.py

Helper functions for building on-brand .pptx decks from a theme.json config.
This is a starting point meant to be adapted per deck, not a black-box CLI.

Requires: python-pptx  (pip install python-pptx --break-system-packages)

Typical usage from a generation script:

    import json
    from apply_theme import load_theme, new_presentation, add_title_slide, \
        add_section_slide, add_content_slide, add_chart_slide

    theme = load_theme("theme.json")
    prs = new_presentation(theme)

    add_title_slide(prs, theme, "Q3 Business Review", "Acme Corp — October 2026")
    add_section_slide(prs, theme, "Performance")
    add_content_slide(prs, theme, "Revenue grew 34% YoY, driven by EMEA",
                       bullets=["EMEA up 61%", "APAC flat", "Churn down 2pts"])
    add_chart_slide(prs, theme, "Monthly active users trending up",
                     categories=["Jan","Feb","Mar","Apr"],
                     series={"MAU": [120, 135, 150, 171]})

    prs.save("/mnt/user-data/outputs/q3_review.pptx")
"""

import json
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE


def load_theme(path):
    with open(path, "r", encoding="utf-8") as f:
        theme = json.load(f)
    return theme


def _rgb(hex_str):
    hex_str = hex_str.lstrip("#")
    return RGBColor.from_string(hex_str.upper())


def new_presentation(theme):
    """Start from the company's own template if available (preserves real
    master slides/layouts), otherwise start blank — caller should then set
    basic slide size etc."""
    template_path = theme.get("template_pptx")
    if template_path and os.path.exists(template_path):
        prs = Presentation(template_path)
        # Strip any pre-existing slides if the template file is just a
        # layout library rather than a deck to extend. Comment out if the
        # template should keep its own starter slides.
        xml_slides = prs.slides._sldIdLst
        slides = list(xml_slides)
        for s in slides:
            xml_slides.remove(s)
    else:
        prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


def _blank_layout(prs):
    # Index 6 is "Blank" in the default PowerPoint template; template decks
    # may differ — adjust if using a custom template_pptx.
    try:
        return prs.slide_layouts[6]
    except IndexError:
        return prs.slide_layouts[-1]


def _add_footer_and_logo(slide, prs, theme, page_number=None):
    rules = theme.get("layout_rules", {})
    logo = theme.get("logo", {})
    colors = theme.get("colors", {})

    if rules.get("footer_text"):
        left, top, width, height = Inches(0.6), prs.slide_height - Inches(0.4), Inches(6), Inches(0.3)
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.text = rules["footer_text"]
        tf.paragraphs[0].font.size = Pt(theme.get("type_scale", {}).get("footnote_pt", 10))
        tf.paragraphs[0].font.color.rgb = _rgb(colors.get("text_primary", "#808080"))

    if rules.get("show_page_numbers") and page_number is not None:
        left = prs.slide_width - Inches(1)
        top = prs.slide_height - Inches(0.4)
        box = slide.shapes.add_textbox(left, top, Inches(0.6), Inches(0.3))
        tf = box.text_frame
        tf.text = str(page_number)
        tf.paragraphs[0].font.size = Pt(theme.get("type_scale", {}).get("footnote_pt", 10))
        tf.paragraphs[0].alignment = PP_ALIGN.RIGHT

    logo_path = logo.get("path")
    if logo_path and os.path.exists(logo_path):
        max_w = Inches(logo.get("max_width_in", 0.9))
        placement = logo.get("placement", "bottom-right")
        if placement == "bottom-right":
            left = prs.slide_width - max_w - Inches(0.3)
            top = prs.slide_height - max_w - Inches(0.2)
        elif placement == "top-right":
            left = prs.slide_width - max_w - Inches(0.3)
            top = Inches(0.2)
        else:  # top-left default fallback
            left, top = Inches(0.3), Inches(0.2)
        slide.shapes.add_picture(logo_path, left, top, width=max_w)


def add_title_slide(prs, theme, title, subtitle=None):
    colors = theme.get("colors", {})
    fonts = theme.get("fonts", {})
    type_scale = theme.get("type_scale", {})

    slide = prs.slides.add_slide(_blank_layout(prs))
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = _rgb(colors.get("primary", "#0B3D91"))

    box = slide.shapes.add_textbox(Inches(0.8), Inches(2.6), prs.slide_width - Inches(1.6), Inches(1.5))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = title
    p = tf.paragraphs[0]
    p.font.size = Pt(type_scale.get("title_slide_pt", 40))
    p.font.bold = True
    p.font.name = fonts.get("heading", "Arial")
    p.font.color.rgb = _rgb(colors.get("text_on_dark", "#FFFFFF"))

    if subtitle:
        sbox = slide.shapes.add_textbox(Inches(0.8), Inches(4.1), prs.slide_width - Inches(1.6), Inches(0.8))
        stf = sbox.text_frame
        stf.text = subtitle
        stf.paragraphs[0].font.size = Pt(18)
        stf.paragraphs[0].font.name = fonts.get("body", "Arial")
        stf.paragraphs[0].font.color.rgb = _rgb(colors.get("text_on_dark", "#FFFFFF"))

    logo = theme.get("logo", {})
    if logo.get("path") and os.path.exists(logo["path"]):
        slide.shapes.add_picture(logo["path"], Inches(0.8), Inches(0.6), width=Inches(logo.get("max_width_in", 0.9)))

    return slide


def add_section_slide(prs, theme, section_title):
    colors = theme.get("colors", {})
    fonts = theme.get("fonts", {})

    slide = prs.slides.add_slide(_blank_layout(prs))
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = _rgb(colors.get("secondary", colors.get("primary", "#0B3D91")))

    box = slide.shapes.add_textbox(Inches(0.8), prs.slide_height / 2 - Inches(0.6), prs.slide_width - Inches(1.6), Inches(1.2))
    tf = box.text_frame
    tf.text = section_title
    p = tf.paragraphs[0]
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.name = fonts.get("heading", "Arial")
    p.font.color.rgb = _rgb(colors.get("text_on_dark", "#FFFFFF"))
    return slide


def add_content_slide(prs, theme, title, bullets=None, page_number=None):
    colors = theme.get("colors", {})
    fonts = theme.get("fonts", {})
    type_scale = theme.get("type_scale", {})
    margins = theme.get("layout_rules", {}).get("margins_in", {"top": 0.5, "left": 0.6, "right": 0.6})

    slide = prs.slides.add_slide(_blank_layout(prs))
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(colors.get("background", "#FFFFFF"))

    tbox = slide.shapes.add_textbox(Inches(margins.get("left", 0.6)), Inches(margins.get("top", 0.5)),
                                     prs.slide_width - Inches(margins.get("left", 0.6) + margins.get("right", 0.6)),
                                     Inches(1))
    ttf = tbox.text_frame
    ttf.word_wrap = True
    ttf.text = title
    ttf.paragraphs[0].font.size = Pt(type_scale.get("slide_title_pt", 28))
    ttf.paragraphs[0].font.bold = True
    ttf.paragraphs[0].font.name = fonts.get("heading", "Arial")
    ttf.paragraphs[0].font.color.rgb = _rgb(colors.get("primary", "#0B3D91"))

    if bullets:
        bbox = slide.shapes.add_textbox(Inches(margins.get("left", 0.6)), Inches(1.7),
                                         prs.slide_width - Inches(margins.get("left", 0.6) + margins.get("right", 0.6)),
                                         Inches(4.5))
        btf = bbox.text_frame
        btf.word_wrap = True
        for i, bullet in enumerate(bullets[:6]):  # enforce 6x6-ish ceiling
            p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
            p.text = f"•  {bullet}"
            p.font.size = Pt(type_scale.get("body_pt", 16))
            p.font.name = fonts.get("body", "Arial")
            p.font.color.rgb = _rgb(colors.get("text_primary", "#1A1A1A"))
            p.space_after = Pt(10)

    _add_footer_and_logo(slide, prs, theme, page_number=page_number)
    return slide


def add_chart_slide(prs, theme, title, categories, series, chart_type="bar", page_number=None):
    """series: dict of {series_name: [values]} aligned with categories."""
    colors = theme.get("colors", {})
    fonts = theme.get("fonts", {})
    type_scale = theme.get("type_scale", {})

    slide = prs.slides.add_slide(_blank_layout(prs))
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(colors.get("background", "#FFFFFF"))

    tbox = slide.shapes.add_textbox(Inches(0.6), Inches(0.5), prs.slide_width - Inches(1.2), Inches(1))
    ttf = tbox.text_frame
    ttf.word_wrap = True
    ttf.text = title
    ttf.paragraphs[0].font.size = Pt(type_scale.get("slide_title_pt", 28))
    ttf.paragraphs[0].font.bold = True
    ttf.paragraphs[0].font.name = fonts.get("heading", "Arial")
    ttf.paragraphs[0].font.color.rgb = _rgb(colors.get("primary", "#0B3D91"))

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series.items():
        chart_data.add_series(name, values)

    xl_type = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    x, y, cx, cy = Inches(0.8), Inches(1.7), prs.slide_width - Inches(1.6), Inches(4.8)
    gframe = slide.shapes.add_chart(xl_type, x, y, cx, cy, chart_data)
    chart = gframe.chart

    chart_sequence = colors.get("chart_sequence", [colors.get("primary", "#0B3D91")])
    for i, plot_series in enumerate(chart.plots[0].series):
        color = chart_sequence[i % len(chart_sequence)]
        plot_series.format.fill.solid()
        plot_series.format.fill.fore_color.rgb = _rgb(color)

    _add_footer_and_logo(slide, prs, theme, page_number=page_number)
    return slide
